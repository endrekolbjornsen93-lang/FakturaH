# Optimera fakturaparser og kategoriserer – Streamlit-versjon
# Kombinerer: PDF-filtrering + parsing + kategorisering + Excel-nedlasting
# Basert på OptimeraFakturaRev02.py

import re
import hashlib
import io
import tempfile
import traceback
from contextlib import redirect_stdout
import warnings
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pdfplumber
from pypdf import PdfReader
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

warnings.filterwarnings('ignore')


# ============================================================================
# DEL 1: PDF-FILTRERING (i minnet - ingen fil skrives)
# ============================================================================

def normalize_text_for_filter(text: str) -> str:
    """Normaliser tekst for side-filtrering."""
    if not text:
        return ""
    return " ".join(text.lower().replace("\xa0", " ").split())


def page_has_side_marker(text: str) -> bool:
    """Sjekk om siden inneholder 'side x av y' mønster."""
    pattern = re.compile(r"side\s*\d+\s*av\s*\d+", flags=re.IGNORECASE)
    return bool(pattern.search(normalize_text_for_filter(text)))


def get_filtered_pages_from_pdf(pdf_path: str) -> List[int]:
    """
    Returner liste over sidenummer (0-indeksert) som skal beholdes.
    Disse sidene har 'side x av y' tekst.
    """
    reader = PdfReader(str(pdf_path))
    filtered_page_indices = []

    for page_num, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        if page_has_side_marker(text):
            filtered_page_indices.append(page_num)

    return filtered_page_indices


# ============================================================================
# DEL 2: FAKTURA-PARSING (fra InvoiceParser-klassen)
# ============================================================================

class InvoiceParser:
    """Parser for norske Optimera-fakturaer/kreditnotaer."""

    def __init__(self, debug: bool = False):
        self.debug = debug
        self.nobb_col_x: Optional[float] = None
        self.art_col_x: Optional[float] = None
        self.benevnelse_col_x: Optional[float] = None
        self.belop_col_x: Optional[float] = None

        self.units = {
            'STK', 'PAK', 'ESK', 'LM', 'M2', 'KG', 'M', 'M3', 'BK', 'L',
            'PAL', 'PLA', 'KRT', 'RUL', 'SPA', 'POS', 'BOX', 'BNT', 'BK'
        }

        self.charge_keywords = (
            'Frakt', 'Express', 'Tidsbestemt', 'Emballasje',
            'RV Frakt', 'Frakt fra leverandør', 'Frakt Rockwool',
            'Tidsvindus', 'Miljøgebyr', 'Miljogebyr', 'Sone', 'Toll', 'Gebyr'
        )

        self.invoice_meta = defaultdict(lambda: {
            'sum_ord': None,
            'fakturatotal': None,
            'charges_anywhere': []
        })

    def parse_pdf_filtered(self, pdf_path: str, page_indices: List[int]) -> List[Dict]:
        """
        Parse kun spesifikke sider fra PDF.
        page_indices: Liste over sidenummer (0-indeksert) som skal parses.
        """
        all_lines: List[Dict] = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                current_faktura: Optional[str] = None
                invoice_line_count = 0

                for page_index in page_indices:
                    if page_index >= len(pdf.pages):
                        continue

                    page = pdf.pages[page_index]
                    text = page.extract_text() or ""

                    # Finn fakturanummer
                    faktura_match = re.search(
                        r'(Faktura(?:nr|nummer)?\.?\s*:?\s*)(\d{5,})',
                        text,
                        flags=re.IGNORECASE
                    )
                    if faktura_match:
                        if self.debug and current_faktura and invoice_line_count > 0:
                            print(f" → Invoice {current_faktura}: {invoice_line_count} lines")
                        current_faktura = faktura_match.group(2)
                        invoice_line_count = 0
                        if self.debug:
                            print(f"[Page {page_index + 1}] Found Faktura {current_faktura}")

                    # Side X av Y
                    side_no, side_total = self._extract_page_no(text)
                    if self.debug and side_no is not None:
                        print(f"[Page {page_index + 1}] Side {side_no} av {side_total}")

                    # Oppdater totalsummer
                    if current_faktura:
                        self._update_invoice_totals(current_faktura, text)
                        self._harvest_charge_candidates(current_faktura, text)

                    # Hent linjer fra siden
                    page_lines = self._extract_lines_from_page(page, current_faktura, side_no=side_no)
                    invoice_line_count += len(page_lines)
                    all_lines.extend(page_lines)

                if self.debug and current_faktura and invoice_line_count > 0:
                    print(f" → Invoice {current_faktura}: {invoice_line_count} lines")

        except Exception as e:
            raise Exception(f"Error parsing PDF: {str(e)}")

        # Avstemming og recovery
        all_lines = self._reconcile_lines_with_totals(all_lines)
        return all_lines

    # [ALLE HJELPEMETODER FRA ORIGINAL INVOICEPARSER - kopieres direkte]
    def _extract_page_no(self, text: str) -> Tuple[Optional[int], Optional[int]]:
        m = re.search(r'Side\s+(\d+)\s+av\s+(\d+)', text, flags=re.IGNORECASE)
        if m:
            try:
                return int(m.group(1)), int(m.group(2))
            except:
                return None, None
        return None, None

    def _norm(self, s: str) -> str:
        if not s:
            return ''
        s = s.lower().strip()
        s = s.replace('ø', 'o').replace('å', 'a').replace('æ', 'ae')
        return re.sub(r'[^a-z0-9]', '', s)

    def _identify_columns(self, words: List[Dict]) -> None:
        self.nobb_col_x = self.art_col_x = self.benevnelse_col_x = self.belop_col_x = None

        def find_word_x(candidates: List[str]) -> Optional[float]:
            cand_norm = {self._norm(c) for c in candidates}
            for w in words:
                if self._norm(w['text']) in cand_norm:
                    return w['x0']
            return None

        self.belop_col_x = find_word_x(["Beløp", "Belop"])
        self.benevnelse_col_x = find_word_x(["Benevnelse", "Beskrivelse", "Varetekst", "Varenavn", "Tekst"])
        # "Aliasnummer" er nytt navn på Nobb-kolonnen i nyere Optimera-fakturaer (fra sept. 2025)
        self.nobb_col_x = find_word_x(["Nobb", "Nobb-nr", "Nobb nr.", "NOBB-nr", "NOBB", "Nobb nr", "Aliasnummer"])
        self.art_col_x = find_word_x(["Art-nr", "Art.nr", "Art.nr.", "Artnr", "Varenr", "Vare nr"])

        if not self.belop_col_x:
            xs = {}
            for w in words:
                if self._is_amount_like_token(w['text']):
                    x = round(w['x0'])
                    xs[x] = xs.get(x, 0) + 1
            if xs:
                self.belop_col_x = max(xs.items(), key=lambda kv: kv[1])[0]

        if not (self.nobb_col_x or self.art_col_x) and words:
            self.art_col_x = words[0]['x0']

    def _group_words_into_rows(self, words: List[Dict]) -> List[Dict]:
        if not words:
            return []
        rows: List[Dict] = []
        current_row = {'words': [words[0]], 'y0': words[0].get('top', words[0].get('y0', 0))}
        for word in words[1:]:
            y = word.get('top', word.get('y0', 0))
            if abs(y - current_row['y0']) < 3:
                current_row['words'].append(word)
            else:
                rows.append(current_row)
                current_row = {'words': [word], 'y0': y}
        if current_row['words']:
            rows.append(current_row)

        for row in rows:
            row['words'].sort(key=lambda w: w['x0'])
        return rows

    def _is_in_belop_column(self, word: Dict) -> bool:
        if not self.belop_col_x:
            return False
        return abs(word['x0'] - self.belop_col_x) < 35

    def _is_amount_like_token(self, text: str) -> bool:
        if not text:
            return False
        t = (text or '').replace('\xa0', ' ').strip()
        t = t.replace('−', '-')
        if t in {'-', '−'}:
            return True
        if t.lower() in {'kr', 'nok'}:
            return True
        if re.match(r'^\d{1,3}$', t):
            return True
        if re.match(r'^\d{1,3}(?:[ .]\d{3})+$', t):
            return True
        if re.match(r'^\d{1,3}[,.]\d{2}-?$', t):
            return True
        if re.match(r'^\d+[,.]\d{2}-?$', t):
            return True
        if re.match(r'^\*\d+$', t):
            return True
        return False

    def _is_amount(self, text: str) -> bool:
        if not text:
            return False
        t = (text or '').replace('\xa0', ' ').strip()
        t = t.replace('−', '-')
        pattern = r'^-?\(?\d{1,3}(?:[ .]\d{3})*\)?[,\.]\d{2}-?$|^-?\(?\d+[,\.]\d{2}\)?-?$'
        return bool(re.match(pattern, t))

    def _clean_amount_token(self, t: str) -> str:
        t = (t or '').replace('\xa0', ' ').strip()
        t = t.replace('−', '-')
        t = t.replace('(', '').replace(')', '')
        t = t.replace('kr', '').replace('KR', '').replace('Kr', '').strip()
        t = re.sub(r'\*\d+', '', t)
        return t

    def _join_amount_parts(self, parts: List[str]) -> Optional[str]:
        if not parts:
            return None
        cleaned = [self._clean_amount_token(p) for p in parts if p and p.strip()]
        candidate = ' '.join(cleaned).strip()
        if self._is_amount(candidate):
            return candidate
        candidate2 = candidate.replace(' ', '')
        if self._is_amount(candidate2):
            return candidate2
        candidate3 = candidate.replace(' ', '.')
        if self._is_amount(candidate3):
            return candidate3
        return None

    def _row_has_amount_in_belop_col(self, row: Dict) -> bool:
        belop_tokens = [w['text'] for w in row['words'] if self._is_in_belop_column(w)]
        if not belop_tokens:
            return False
        joined = self._join_amount_parts(belop_tokens)
        return bool(joined and self._is_amount(joined))

    def _is_skip_row(self, row: Dict) -> bool:
        row_text = ' '.join(w['text'] for w in row['words'])
        skip_keywords = [
            'Ordretotal', 'Fakturatotal', 'Sum ordrelinjer', 'Sum å betale',
            'Selger har', 'Leveringsmåte', 'Deres referanse', 'Vår referanse',
            'Ordrenummer', 'Utleveringssted', 'Kundens ordrenr', 'Leveringsinformasjon',
            'NB NB', 'OMSTILLINGSTILLEGG', 'MVA', 'Avrunding', 'Kid-nummer', 'KID-nummer',
            'Fakturabeløp'
        ]
        return any(keyword in row_text for keyword in skip_keywords)

    def _is_numeric_row(self, row: Dict) -> bool:
        texts = [w['text'] for w in row['words']]
        if not texts:
            return False
        numeric_count = sum(1 for t in texts if self._is_number(t) or t in self.units)
        return numeric_count / len(texts) > 0.6

    def _is_number(self, text: str) -> bool:
        cleaned = (text or '').replace('\xa0', ' ')
        cleaned = cleaned.replace('−', '-')
        cleaned = cleaned.replace(' ', '').replace('.', '').replace(',', '.')
        try:
            float(cleaned)
            return True
        except ValueError:
            return False

    def _is_art_candidate(self, text: str) -> bool:
        t = (text or '').strip()
        return (t.isdigit() and 6 <= len(t) <= 10) or bool(re.match(r'^[A-Za-z]\w{5,}$', t))

    def _find_charge_keyword_in_row(self, row: Dict) -> Optional[str]:
        row_text = ' '.join(w['text'] for w in row['words'])
        for k in self.charge_keywords:
            if re.search(rf'\b{re.escape(k)}\b', row_text, flags=re.IGNORECASE):
                return k
        return None

    def _find_amount_anywhere_in_row(self, row: Dict) -> Optional[str]:
        candidates = []
        prev_minus = False
        for w in row['words']:
            t_raw = w['text']
            t = self._clean_amount_token(t_raw)
            if t_raw in {'-', '−'}:
                prev_minus = True
                continue
            if self._is_amount(t):
                candidates.append((w.get('x0', 0.0), t))
                prev_minus = False
                continue
            if prev_minus and self._is_amount('-' + t):
                candidates.append((w.get('x0', 0.0), '-' + t))
            prev_minus = False

        if not candidates:
            row_text = ' '.join(w['text'] for w in row['words'])
            row_text = row_text.replace('\xa0', ' ').replace('−', '-')
            row_text = re.sub(r'\b[Kk][Rr]\b', '', row_text)
            m_all = re.findall(r'-?\(?\d{1,3}(?:[ .]\d{3})*\)?[,\.]\d{2}-?', row_text)
            if m_all:
                return m_all[-1]

        if candidates:
            if self.belop_col_x is not None:
                candidates.sort(key=lambda it: abs(it[0] - self.belop_col_x))
            else:
                candidates.sort(key=lambda it: it[0], reverse=True)
            return candidates[0][1]
        return None

    def _extract_lines_from_page(self, page, current_faktura: Optional[str], side_no: Optional[int]) -> List[Dict]:
        lines: List[Dict] = []
        words = page.extract_words() or []
        if not words or not current_faktura:
            return lines

        self._identify_columns(words)
        if self.debug:
            print(
                f"  Columns x: NOBB={self.nobb_col_x} | ART={self.art_col_x} | DESC={self.benevnelse_col_x} | BELØP={self.belop_col_x}")

        if not self.belop_col_x:
            return lines

        rows = self._group_words_into_rows(words)

        table_start_y = None
        table_end_y = None

        header_triggers = ("Benevnelse", "Beskrivelse", "Varetekst", "Varenavn", "Tekst")
        footer_triggers = (
            "Fakturatotal", "Sum ordrelinjer", "Sum å betale",
            "MVA %", "MVA", "Kid-nummer", "KID-nummer",
            "Fakturabeløp", "Avrunding", "Selger har salgspant"
        )

        for row in rows:
            row_text = ' '.join(w['text'] for w in row['words'])
            if any(h in row_text for h in header_triggers):
                table_start_y = row['y0']
                break

        if table_start_y is not None:
            for row in rows:
                if row['y0'] <= table_start_y:
                    continue
                row_text = ' '.join(w['text'] for w in row['words'])
                if any(k in row_text for k in footer_triggers):
                    table_end_y = row['y0']
                    break

        if table_start_y is None:
            if side_no is not None and side_no > 1:
                for row in rows:
                    if self._row_has_amount_in_belop_col(row) and not self._is_skip_row(row):
                        table_start_y = row['y0'] - 2
                        break
            if table_start_y is None:
                for row in rows:
                    if self._row_has_amount_in_belop_col(row) and not self._is_skip_row(row):
                        table_start_y = row['y0'] - 2
                        break

        if table_start_y is None:
            return lines
        if table_end_y is None:
            table_end_y = page.height

        table_rows = [r for r in rows if table_start_y < r['y0'] < table_end_y]

        if self.debug:
            print(f"  Table area rows: {len(table_rows)} (y:{table_start_y:.1f}→{table_end_y:.1f})")

        i = 0
        while i < len(table_rows):
            row = table_rows[i]
            has_belop = self._row_has_amount_in_belop_col(row)
            is_charge = self._find_charge_keyword_in_row(row) is not None

            if (has_belop or is_charge) and not self._is_skip_row(row):
                parsed, rows_consumed = self._parse_product_from_rows(table_rows, i, current_faktura)
                if parsed:
                    lines.append(parsed)
                    if self.debug:
                        print(
                            f"   Extracted: {parsed.get('Beskrivelse', '')[:80]}  | Artnr={parsed.get('Art-nr', '')} | Beløp={parsed.get('Beløp', '')}")
                i += rows_consumed
            else:
                i += 1

        return lines

    def _find_artnr_below(self, rows: List[Dict], start_idx: int) -> Optional[str]:
        cols_to_check: List[float] = []
        if self.art_col_x is not None:
            cols_to_check.append(self.art_col_x)
        leftmost_col_x = self.nobb_col_x if self.nobb_col_x is not None else self.art_col_x
        if leftmost_col_x is not None and (not cols_to_check or leftmost_col_x != cols_to_check[0]):
            cols_to_check.append(leftmost_col_x)

        for j in range(start_idx + 1, min(start_idx + 4, len(rows))):
            row = rows[j]
            for w in row['words']:
                for colx in cols_to_check:
                    if abs(w['x0'] - colx) < 20 and self._is_art_candidate(w['text']):
                        return w['text']
        return None

    def _row_contains_charge_keyword(self, row: Dict) -> bool:
        row_text = ' '.join(w['text'] for w in row['words'])
        return any(k in row_text for k in self.charge_keywords)

    def _parse_product_from_rows(self, rows: List[Dict], start_idx: int, faktura: str) -> Tuple[Optional[Dict], int]:
        nobb_nr = ''
        art_nr = ''
        beskrivelse_parts: List[str] = []
        belop = ''
        rows_consumed = 1

        leftmost_col_x = self.nobb_col_x if self.nobb_col_x is not None else self.art_col_x
        first_row = rows[start_idx]

        charge_kw = self._find_charge_keyword_in_row(first_row)

        belop_tokens = []
        for w in first_row['words']:
            if self._is_in_belop_column(w):
                belop_tokens.append(w['text'])
        belop_joined = self._join_amount_parts(belop_tokens)
        if belop_joined:
            belop = belop_joined

        if not belop and charge_kw:
            belop_any = self._find_amount_anywhere_in_row(first_row)
            if belop_any:
                belop = belop_any

        for w in first_row['words']:
            x0 = w['x0']
            text = w['text']

            if charge_kw and re.fullmatch(re.escape(charge_kw), text, flags=re.IGNORECASE):
                if charge_kw.lower() not in (bp.lower() for bp in beskrivelse_parts):
                    beskrivelse_parts.append(charge_kw)

            if leftmost_col_x is not None and abs(x0 - leftmost_col_x) < 15:
                if text.isdigit() and len(text) == 8:
                    nobb_nr = text
                    continue

            if self.art_col_x is not None and abs(x0 - self.art_col_x) < 20 and self._is_art_candidate(text):
                art_nr = text
                continue

            if self.benevnelse_col_x is not None and x0 >= self.benevnelse_col_x - 10:
                if not self._is_in_belop_column(w):
                    if (text in self.units or self._is_number(text)) and not self._row_contains_charge_keyword(
                            first_row):
                        continue
                    beskrivelse_parts.append(text)

        if charge_kw and not beskrivelse_parts:
            beskrivelse_parts = [charge_kw]

        if not belop:
            return None, 1

        if not art_nr:
            art_nr = self._find_artnr_below(rows, start_idx) or ''

        for i in range(start_idx + 1, len(rows)):
            row = rows[i]

            if self._row_has_amount_in_belop_col(row):
                break

            if not art_nr and self.art_col_x is not None:
                for w in row['words']:
                    if abs(w['x0'] - self.art_col_x) < 20 and self._is_art_candidate(w['text']):
                        art_nr = w['text']
                        break

            if not art_nr and leftmost_col_x is not None:
                for w in row['words']:
                    if abs(w['x0'] - leftmost_col_x) < 15 and self._is_art_candidate(w['text']):
                        art_nr = w['text']
                        break

            if self._is_skip_row(row):
                break

            rows_consumed += 1

            for w in row['words']:
                x0 = w['x0']
                text = w['text']

                if self._is_in_belop_column(w):
                    continue

                if (leftmost_col_x is not None and abs(x0 - leftmost_col_x) < 15) or \
                        (self.art_col_x is not None and abs(x0 - self.art_col_x) < 20 and self._is_art_candidate(text)):
                    continue

                if (text in self.units or self._is_number(text)) and not self._row_contains_charge_keyword(row):
                    continue

                if self.benevnelse_col_x is None or x0 >= self.benevnelse_col_x - 10:
                    beskrivelse_parts.append(text)

            if charge_kw and not belop:
                belop_any = self._find_amount_anywhere_in_row(row)
                if belop_any:
                    belop = belop_any

            if self._is_numeric_row(row):
                break
            if rows_consumed >= 8:
                break

        beskrivelse = ' '.join(beskrivelse_parts).strip()

        if belop and (beskrivelse or art_nr or nobb_nr):
            return {
                'Fakturanr': faktura,
                'Nobb-nr': nobb_nr,
                'Art-nr': art_nr,
                'Beskrivelse': beskrivelse,
                'Beløp': belop
            }, rows_consumed

        return None, rows_consumed

    def _parse_no_amount(self, s: str) -> Optional[float]:
        if not s:
            return None
        t = s.replace('\xa0', ' ').strip()
        t = t.replace('−', '-')
        t = t.replace('kr', '').replace('KR', '').replace('Kr', '').strip()
        t = t.replace(' ', '').replace('.', '')
        if t.endswith('-') and '-' not in t[:-1]:
            t = '-' + t[:-1]
        paren_neg = t.startswith('(') and t.endswith(')')
        t = t.replace('(', '').replace(')', '')
        t = t.replace(',', '.')
        try:
            val = float(t)
            if paren_neg and val > 0:
                val = -val
            return val
        except:
            return None

    def _update_invoice_totals(self, fakturanr: str, text: str) -> None:
        if not fakturanr or not text:
            return
        t = text.replace('\xa0', ' ').replace('−', '-')

        m1 = re.search(r'Sum\s+ordrelinjer\s*:?\s*(-?\(?\d{1,3}(?:[ .]\d{3})*\)?[,\.]\d{2})', t, flags=re.IGNORECASE)
        if m1:
            v = self._parse_no_amount(m1.group(1))
            if v is not None:
                self.invoice_meta[fakturanr]['sum_ord'] = v

        m2 = re.search(r'Fakturatotal\s*:?\s*(-?\(?\d{1,3}(?:[ .]\d{3})*\)?[,\.]\d{2})', t, flags=re.IGNORECASE)
        if m2:
            v = self._parse_no_amount(m2.group(1))
            if v is not None:
                self.invoice_meta[fakturanr]['fakturatotal'] = v

    def _harvest_charge_candidates(self, fakturanr: str, text: str) -> None:
        if not fakturanr or not text:
            return
        t = text.replace('\xa0', ' ').replace('−', '-')
        lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
        kw_pattern = r'(Frakt|Express|Emballasje|Tidsbestemt|Tidsvindus|Miljøgebyr|Miljogebyr|Sone|Toll|Gebyr)'
        amt_pattern = r'-?\(?\d{1,3}(?:[ .]\d{3})*\)?[,\.]\d{2}-?'
        for ln in lines:
            if re.search(r'(Sum ordrelinjer|Fakturatotal|Fakturabeløp|MVA|Avrunding|KID-?nummer)', ln,
                         flags=re.IGNORECASE):
                continue
            m_kw = re.search(kw_pattern, ln, flags=re.IGNORECASE)
            m_amt = re.search(amt_pattern, ln)
            if m_kw and m_amt:
                desc = m_kw.group(1)
                belop = m_amt.group(0)
                self.invoice_meta[fakturanr]['charges_anywhere'].append({'desc': desc, 'belop': belop})

    def _reconcile_lines_with_totals(self, all_lines: List[Dict]) -> List[Dict]:
        if not all_lines:
            return all_lines

        by_inv: Dict[str, List[Dict]] = {}
        for d in all_lines:
            by_inv.setdefault(d.get('Fakturanr', ''), []).append(d)

        reconciled: List[Dict] = []
        for faktura, rows in by_inv.items():
            rows_out = list(rows)
            meta = self.invoice_meta.get(faktura, {})
            target = meta.get('sum_ord', None)
            if target is None:
                target = meta.get('fakturatotal', None)

            s = 0.0
            for r in rows_out:
                v = self._to_float_no(r.get('Beløp'))
                if v is not None:
                    s += v

            if target is None:
                reconciled.extend(rows_out)
                continue

            diff = round(target - s, 2)
            if abs(diff) <= 0.01:
                reconciled.extend(rows_out)
                continue

            candidates = meta.get('charges_anywhere', []) or []
            existing = {(r.get('Beskrivelse', '').lower(), round(self._to_float_no(r.get('Beløp')) or 0.0, 2)) for r in
                        rows_out}

            def norm(x: str) -> str:
                return (x or '').strip().lower()

            to_add = []
            for c in candidates:
                c_amt = self._parse_no_amount(c['belop'])
                if c_amt is None:
                    continue
                key_opts = [(norm(c['desc']), round(c_amt, 2)), ('frakt', round(c_amt, 2))]
                if any(k in existing for k in key_opts):
                    continue
                to_add.append({'Fakturanr': faktura,
                               'Nobb-nr': '',
                               'Art-nr': '',
                               'Beskrivelse': c['desc'],
                               'Beløp': c['belop']})

            for x in to_add:
                rows_out.append(x)
                s += self._to_float_no(x['Beløp']) or 0.0
                if abs(round(target - s, 2)) <= 0.01:
                    break
            reconciled.extend(rows_out)

            if self.debug:
                final_diff = round(target - s, 2)
                if abs(final_diff) > 0.01:
                    print(f"⚠ Avstemming mislyktes for faktura {faktura}: diff {final_diff} etter recovery")

        return reconciled

    def _to_float_no(self, belop_str: str) -> Optional[float]:
        if not belop_str or not isinstance(belop_str, str):
            return None
        raw = belop_str
        t = belop_str.replace('\xa0', ' ').strip()
        t = t.replace('−', '-')
        paren_neg = ('(' in raw and ')' in raw)

        t = t.replace('kr', '').replace('KR', '').replace('Kr', '').strip()
        t = t.replace(' ', '').replace('.', '')
        if t.endswith('-') and '-' not in t[:-1]:
            t = '-' + t[:-1]
        t = t.replace('(', '').replace(')', '')
        t = t.replace(',', '.')
        try:
            val = float(t)
            if paren_neg and val > 0:
                val = -val
            return val
        except ValueError:
            return None


# ============================================================================
# DEL 3: KATEGORISERING (basert på Fakturasplitter)
# ============================================================================

def normaliser_id(ser):
    """Kast til str, trim, fjern mellomrom. Håndter tall riktig."""

    def clean_val(val):
        if pd.isna(val):
            return ""
        # Hvis det er et tall (int eller float), konverter til int først
        if isinstance(val, (int, float)):
            # Fjern desimaler hvis det er et heltall
            if val == int(val):
                return str(int(val)).strip().replace(" ", "")
            else:
                return str(val).strip().replace(" ", "")
        # Ellers vanlig string-behandling
        return str(val).strip().replace(" ", "")

    return ser.apply(clean_val)


def serie_tom(ser):
    """Sjekk om serie er tom/NaN/blank."""
    return ser.isna() | (ser.astype(str).str.strip() == "") | (
        ser.astype(str).str.strip().str.lower().isin(["nan", "none"]))


def finn_kolonnenavn(df, søk):
    """finn kolonne som inneholder søkeord (case-insensitive)."""
    søk = søk.lower()
    for kol in df.columns:
        if søk in str(kol).lower().strip():
            return kol
    return None


def kategoriser_og_eksporter(df_varelinjer_parsed: pd.DataFrame, varelinjer_kategorisert_fil: str, output_file: str,
                             debug: bool = False):
    """
    Tar inn parsed varelinjer (DataFrame) og varelinjer_kategorisert (Excel-fil),
    kobler dem sammen og eksporterer til output_file.
    """

    # Les varelinjer_kategorisert - FANEN "Linjer"
    # dtype=str: les alt som tekst. Uten dette gjør pandas en kolonne med tomme celler om til
    # desimaltall (21258439 -> "21258439.0"), og da matcher ikke Nobbnr/Artnr mot fakturaen.
    df_varelinjer_kat = pd.read_excel(varelinjer_kategorisert_fil, sheet_name="Linjer", engine="openpyxl", dtype=str)
    print(f"  ✓ Leste {len(df_varelinjer_kat)} rader fra 'Linjer'-fanen")

    # --- PATCH: tving alle kolonner til tekst (eliminerer float/NaN i EXE) ---
    for _col in df_varelinjer_kat.columns:
        df_varelinjer_kat[_col] = df_varelinjer_kat[_col].astype(str).replace("nan", "").fillna("")

    # Les også "Varer med Variasjon" fanen
    try:
        df_flere_kat = pd.read_excel(varelinjer_kategorisert_fil, sheet_name="Varer med Variasjon", engine="openpyxl", dtype=str)
        print(f"  ✓ Leste {len(df_flere_kat)} rader fra 'Varer med Variasjon'")

        # --- PATCH: tving også her alle kolonner til tekst ---
        if not df_flere_kat.empty:
            for _col in df_flere_kat.columns:
                df_flere_kat[_col] = df_flere_kat[_col].astype(str).replace("nan", "").fillna("")

    except Exception as e:
        print(f"  ⚠ Kunne ikke lese 'Varer med Variasjon': {e}")
        df_flere_kat = pd.DataFrame()  # Tom hvis fanen ikke finnes

    # Trim kolonnenavn
    df_varelinjer_parsed.columns = df_varelinjer_parsed.columns.map(lambda x: str(x).strip())

    # --- PATCH: tving alle parsed-kolonner til tekst ---
    for _col in df_varelinjer_parsed.columns:
        df_varelinjer_parsed[_col] = df_varelinjer_parsed[_col].astype(str).replace("nan", "").fillna("")

    df_varelinjer_kat.columns = df_varelinjer_kat.columns.map(lambda x: str(x).strip())

    # Finn kolonner i parsed data
    kol_fakturanr = finn_kolonnenavn(df_varelinjer_parsed, "fakturanr")
    kol_beløp = finn_kolonnenavn(df_varelinjer_parsed, "beløp")
    kol_beskrivelse = finn_kolonnenavn(df_varelinjer_parsed, "beskrivelse") or "Beskrivelse"
    kol_nobbnr = finn_kolonnenavn(df_varelinjer_parsed, "nobbnr")
    kol_artnr = finn_kolonnenavn(df_varelinjer_parsed, "artnr")

    if kol_fakturanr is None or kol_beløp is None:
        raise ValueError("Fant ikke påkrevde kolonner: 'fakturanr' og 'beløp' må finnes.")

    if kol_beskrivelse not in df_varelinjer_parsed.columns:
        df_varelinjer_parsed[kol_beskrivelse] = ""

    # Normaliser ID-kolonner
    if kol_nobbnr and kol_nobbnr in df_varelinjer_parsed.columns:
        df_varelinjer_parsed[kol_nobbnr] = normaliser_id(df_varelinjer_parsed[kol_nobbnr].astype(str))
    if kol_artnr and kol_artnr in df_varelinjer_parsed.columns:
        df_varelinjer_parsed[kol_artnr] = normaliser_id(df_varelinjer_parsed[kol_artnr].astype(str))

    # Behold KUN varer med kode
    if kol_nobbnr or kol_artnr:
        mask_kode = pd.Series([False] * len(df_varelinjer_parsed), index=df_varelinjer_parsed.index)
        if kol_nobbnr:
            mask_kode = mask_kode | (~serie_tom(df_varelinjer_parsed[kol_nobbnr]))
        if kol_artnr:
            mask_kode = mask_kode | (~serie_tom(df_varelinjer_parsed[kol_artnr]))
        df_varelinjer_parsed = df_varelinjer_parsed[mask_kode].copy()

    # Finn kolonner i varelinjer_kategorisert
    kol_vare_nobbnr = finn_kolonnenavn(df_varelinjer_kat, "nobbnr")
    kol_vare_artnr = finn_kolonnenavn(df_varelinjer_kat, "artnr")
    kol_spesifikasjon = finn_kolonnenavn(df_varelinjer_kat, "spesifikasjon")
    kol_fag = finn_kolonnenavn(df_varelinjer_kat, "fag")
    kol_område = finn_kolonnenavn(df_varelinjer_kat, "område")
    kol_frikode = finn_kolonnenavn(df_varelinjer_kat, "frikode")
    kol_vare_fakturanr = finn_kolonnenavn(df_varelinjer_kat, "fakturanummer")
    kol_vare_beskrivelse = finn_kolonnenavn(df_varelinjer_kat, "beskrivelse")

    # Normaliser ID-kolonner i varelinjer_kat
    if kol_vare_nobbnr and kol_vare_nobbnr in df_varelinjer_kat.columns:
        df_varelinjer_kat[kol_vare_nobbnr] = normaliser_id(df_varelinjer_kat[kol_vare_nobbnr].astype(str))
    if kol_vare_artnr and kol_vare_artnr in df_varelinjer_kat.columns:
        df_varelinjer_kat[kol_vare_artnr] = normaliser_id(df_varelinjer_kat[kol_vare_artnr].astype(str))

    kategori_kolonner = ["Spesifikasjon", "Fag", "Område", "Frikode"]

    # Trekk kategorier inn på parsed data
    def hent_kategorier(rad):
        nobb = rad.get(kol_nobbnr) if kol_nobbnr else None
        art = rad.get(kol_artnr) if kol_artnr else None
        besk = rad.get(kol_beskrivelse) if kol_beskrivelse else None
        tom = pd.Series([None] * 4, index=kategori_kolonner)

        # Normaliser verdier for matching - håndter tall riktig
        def normalize_id(val):
            if pd.isna(val):
                return None
            if isinstance(val, (int, float)):
                if val == int(val):
                    return str(int(val)).strip().replace(" ", "")
                else:
                    return str(val).strip().replace(" ", "")
            return str(val).strip().replace(" ", "")

        nobb_norm = normalize_id(nobb)
        art_norm = normalize_id(art)

        # Prioriter Nobbnr først i "Linjer"
        if nobb_norm and kol_vare_nobbnr:
            treff = df_varelinjer_kat[df_varelinjer_kat[kol_vare_nobbnr] == nobb_norm]
            if not treff.empty:
                data = treff[[kol_spesifikasjon, kol_fag, kol_område, kol_frikode]].iloc[0]
                data.index = tom.index
                return data

        # Hvis ingen Nobbnr-treff, prøv Artnr i "Linjer"
        if art_norm and kol_vare_artnr:
            treff = df_varelinjer_kat[df_varelinjer_kat[kol_vare_artnr] == art_norm]
            if not treff.empty:
                data = treff[[kol_spesifikasjon, kol_fag, kol_område, kol_frikode]].iloc[0]
                data.index = tom.index
                return data

        # Hvis INGEN match i "Linjer", prøv eksakt beskrivelsesmatch i "Varer med Variasjon"
        if besk and pd.notna(besk) and not df_flere_kat.empty:
            besk_norm = str(besk).strip()
            kol_fler_besk = finn_kolonnenavn(df_flere_kat, "beskrivelse")
            if kol_fler_besk:
                treff_besk = df_flere_kat[df_flere_kat[kol_fler_besk].astype(str).str.strip() == besk_norm]
                if not treff_besk.empty:
                    # Finn kategorikolonner i df_flere_kat
                    kol_fler_spes = finn_kolonnenavn(df_flere_kat, "spesifikasjon")
                    kol_fler_fag = finn_kolonnenavn(df_flere_kat, "fag")
                    kol_fler_omr = finn_kolonnenavn(df_flere_kat, "område")
                    kol_fler_fri = finn_kolonnenavn(df_flere_kat, "frikode")

                    if all([kol_fler_spes, kol_fler_fag, kol_fler_omr, kol_fler_fri]):
                        data = treff_besk[[kol_fler_spes, kol_fler_fag, kol_fler_omr, kol_fler_fri]].iloc[0]
                        data.index = tom.index
                        return data

        return tom

    df_varelinjer_parsed[kategori_kolonner] = df_varelinjer_parsed.apply(hent_kategorier, axis=1)

    # Debug: Vis hvor mange som fikk kategorier
    antall_med_kat = df_varelinjer_parsed[kategori_kolonner].notna().any(axis=1).sum()
    antall_total = len(df_varelinjer_parsed)
    print(f"  ✓ {antall_med_kat} av {antall_total} varelinjer fikk kategorier")
    print(f"  → {antall_total - antall_med_kat} varelinjer uten kategorier")

    # Bygg Vareoversikt
    vareoversikt = pd.DataFrame({
        "Fakturanr": df_varelinjer_parsed[kol_fakturanr],
        "Spesifikasjon": df_varelinjer_parsed["Spesifikasjon"],
        "Fag": df_varelinjer_parsed["Fag"],
        "Område": df_varelinjer_parsed["Område"],
        "Frikode": df_varelinjer_parsed["Frikode"],
        "Beskrivelse": df_varelinjer_parsed[kol_beskrivelse],
        "Beløp": pd.to_numeric(df_varelinjer_parsed[kol_beløp].apply(
            lambda x: InvoiceParser()._to_float_no(x) if isinstance(x, str) else x), errors="coerce"),
        "Nobbnr": df_varelinjer_parsed[kol_nobbnr] if kol_nobbnr in df_varelinjer_parsed.columns else "",
        "Artnr": df_varelinjer_parsed[kol_artnr] if kol_artnr in df_varelinjer_parsed.columns else "",
    })

    # Ukjente varer - varer UTEN match (alle 4 kategorier tomme/None)
    def er_tom(val):
        """Sjekk om verdi er tom (None, NaN, tom string, 'nan' som string)"""
        if pd.isna(val):
            return True
        if isinstance(val, str) and val.strip() in ['', 'nan', 'None']:
            return True
        return False

    ukjente_mask = (
            vareoversikt['Spesifikasjon'].apply(er_tom) &
            vareoversikt['Fag'].apply(er_tom) &
            vareoversikt['Område'].apply(er_tom) &
            vareoversikt['Frikode'].apply(er_tom)
    )

    ukjente_varer = vareoversikt[ukjente_mask].copy()

    # Bare ta med relevante kolonner
    ukjente_varer = ukjente_varer[["Fakturanr", "Beskrivelse", "Beløp", "Nobbnr", "Artnr"]].copy()

    if len(ukjente_varer) > 0:
        print(f"  ℹ {len(ukjente_varer)} ukjente varer (ingen kategorier)")
    else:
        print(f"  ✓ Alle varer fikk kategorier")

    # Definer desired_cols tidlig
    desired_cols_var = ["Fakturanr", "Beløp", "Nobbnr", "Artnr", "Spesifikasjon", "Fag", "Område", "Frikode",
                        "Beskrivelse"]

    # Debug: Vis hva vi fikk i Vareoversikt
    if debug:
        print(f"  Debug: Vareoversikt kolonner: {vareoversikt.columns.tolist()}")
        print(f"  Debug: Første 3 Nobbnr fra Vareoversikt: {vareoversikt['Nobbnr'].head(3).tolist()}")

    # Varer med variasjon - vis fakturanr, beløp og alle unike kombinasjoner
    if not df_flere_kat.empty:
        df_flere_kat.columns = df_flere_kat.columns.map(lambda x: str(x).strip())

        # Normaliser ID-kolonner
        kol_fler_nobbnr = finn_kolonnenavn(df_flere_kat, "nobbnr")
        kol_fler_artnr = finn_kolonnenavn(df_flere_kat, "artnr")
        kol_fler_spes = finn_kolonnenavn(df_flere_kat, "spesifikasjon")
        kol_fler_fag = finn_kolonnenavn(df_flere_kat, "fag")
        kol_fler_omr = finn_kolonnenavn(df_flere_kat, "område")
        kol_fler_fri = finn_kolonnenavn(df_flere_kat, "frikode")
        kol_fler_besk = finn_kolonnenavn(df_flere_kat, "beskrivelse")

        # Debug: Vis hva vi fant
        if debug:
            print(f"  Debug: Kolonne i 'Varer med Variasjon': Nobbnr={kol_fler_nobbnr}, Artnr={kol_fler_artnr}")

        if kol_fler_nobbnr and kol_fler_nobbnr in df_flere_kat.columns:
            df_flere_kat[kol_fler_nobbnr] = normaliser_id(df_flere_kat[kol_fler_nobbnr])
            if debug:
                print(
                    f"  Debug: Første 5 Nobbnr fra 'Varer med Variasjon': {df_flere_kat[kol_fler_nobbnr].head().tolist()}")
        if kol_fler_artnr and kol_fler_artnr in df_flere_kat.columns:
            df_flere_kat[kol_fler_artnr] = normaliser_id(df_flere_kat[kol_fler_artnr])

        # Samle alle rader som kombinerer fakturadata med variasjonskombinasjonene
        var_resultat_rows = []

        def normalize_id_for_match(val):
            if pd.isna(val):
                return None
            if isinstance(val, (int, float)):
                if val == int(val):
                    return str(int(val)).strip()
                else:
                    return str(val).strip()
            return str(val).strip()

        # Debug: Tell hvor mange vi finner
        match_count = 0
        total_checked = 0

        for _, rad in vareoversikt.iterrows():
            faktura = rad['Fakturanr']
            nobb_raw = rad.get('Nobbnr')
            art_raw = rad.get('Artnr')
            nobb = normalize_id_for_match(nobb_raw)
            art = normalize_id_for_match(art_raw)
            belop = rad['Beløp']
            besk_faktura = rad['Beskrivelse']

            # Debug første rad
            if debug and total_checked == 0:
                print(
                    f"  Debug: Første vare fra faktura - Nobbnr_raw={nobb_raw} (type={type(nobb_raw)}), normalized={nobb}")

            total_checked += 1

            # Sjekk om Nobbnr finnes i "Varer med Variasjon"
            if nobb and nobb not in ['', 'nan', 'None'] and kol_fler_nobbnr:
                treff = df_flere_kat[df_flere_kat[kol_fler_nobbnr] == nobb]

                if debug and total_checked <= 3:
                    print(f"  Debug: Søker etter Nobbnr='{nobb}' i 'Varer med Variasjon' - fant {len(treff)} treff")

                if not treff.empty:
                    match_count += 1
                    # Hent alle unike kombinasjoner
                    for _, kombo in treff.iterrows():
                        var_resultat_rows.append({
                            'Fakturanr': faktura,
                            'Beløp': belop,
                            'Nobbnr': nobb,
                            'Artnr': art if art else '',
                            'Spesifikasjon': kombo.get(kol_fler_spes, ''),
                            'Fag': kombo.get(kol_fler_fag, ''),
                            'Område': kombo.get(kol_fler_omr, ''),
                            'Frikode': kombo.get(kol_fler_fri, ''),
                            'Beskrivelse': kombo.get(kol_fler_besk, '') if kol_fler_besk else besk_faktura
                        })
                    continue

            # Hvis ingen Nobbnr, sjekk Artnr
            if art and art not in ['', 'nan', 'None'] and kol_fler_artnr:
                if not nobb or nobb in ['', 'nan', 'None']:
                    treff = df_flere_kat[df_flere_kat[kol_fler_artnr] == art]
                    if not treff.empty:
                        match_count += 1
                        # Hent alle unike kombinasjoner
                        for _, kombo in treff.iterrows():
                            var_resultat_rows.append({
                                'Fakturanr': faktura,
                                'Beløp': belop,
                                'Nobbnr': '',
                                'Artnr': art,
                                'Spesifikasjon': kombo.get(kol_fler_spes, ''),
                                'Fag': kombo.get(kol_fler_fag, ''),
                                'Område': kombo.get(kol_fler_omr, ''),
                                'Frikode': kombo.get(kol_fler_fri, ''),
                                'Beskrivelse': kombo.get(kol_fler_besk, '') if kol_fler_besk else besk_faktura
                            })

        print(f"  → Sjekket {total_checked} varer fra faktura, fant {match_count} med variasjon")

        # Drop duplikater og sorter
        if var_resultat_rows:
            nobb_resultat = pd.DataFrame(var_resultat_rows).drop_duplicates()
            nobb_resultat = nobb_resultat[desired_cols_var].sort_values(
                by=["Fakturanr", "Nobbnr", "Artnr", "Spesifikasjon", "Fag", "Område", "Frikode"]
            ).reset_index(drop=True)

            # Tell unike varer og kombinasjoner
            unike_varer = nobb_resultat[['Nobbnr', 'Artnr']].drop_duplicates()
            print(f"  ✓ {len(nobb_resultat)} rader ({len(unike_varer)} unike varer med variasjon)")
        else:
            nobb_resultat = pd.DataFrame(columns=desired_cols_var)
            print(f"  ℹ Ingen varer med variasjon funnet i fakturaen")
    else:
        nobb_resultat = pd.DataFrame(columns=desired_cols_var)
        print(f"  ⚠ 'Varer med Variasjon'-fanen er tom")

    # Sett opp mengder for markering i Vareoversikt
    # Marker varer som finnes i "Varer med flere kategorier"
    if not nobb_resultat.empty:
        nobb_var_set = set(
            nobb_resultat.loc[nobb_resultat["Nobbnr"].astype(str).str.strip() != "", "Nobbnr"].astype(str))
        art_var_set = set(nobb_resultat.loc[(nobb_resultat["Nobbnr"].astype(str).str.strip() == "") &
                                            (nobb_resultat["Artnr"].astype(str).str.strip() != ""), "Artnr"].astype(
            str))
    else:
        nobb_var_set = set()
        art_var_set = set()

    # Marker rader som skal være GULE
    # a) Nobbnr i variasjonsliste
    nobb_i_variasjon = vareoversikt["Nobbnr"].astype(str).isin(nobb_var_set)

    # b) Artnr i variasjonsliste når Nobbnr mangler på raden
    nobb_mangler = serie_tom(vareoversikt["Nobbnr"])
    art_i_variasjon_uten_nobb = vareoversikt["Artnr"].astype(str).isin(art_var_set) & nobb_mangler

    # Total flagg
    flag_vare_gul = nobb_i_variasjon | art_i_variasjon_uten_nobb

    # Kategorisert (unike kombinasjoner med summert beløp)
    kategorisert_temp = vareoversikt.dropna(subset=["Spesifikasjon", "Fag", "Område", "Frikode"]).copy()

    if not kategorisert_temp.empty:
        kategorisert = (
            kategorisert_temp
            .groupby(["Fakturanr", "Spesifikasjon", "Fag", "Område", "Frikode"], as_index=False)
            .agg({"Beløp": "sum"})
            .rename(columns={"Beløp": "Summert Beløp"})
            .sort_values(by=["Fakturanr", "Spesifikasjon", "Fag", "Område", "Frikode"])
            .reset_index(drop=True)
        )
    else:
        kategorisert = pd.DataFrame(columns=["Fakturanr", "Spesifikasjon", "Fag", "Område", "Frikode", "Summert Beløp"])

    # --- PATCH: 100% robust KEY-join som tåler float/NaN/None ---
    key_cols = ["Fakturanr", "Spesifikasjon", "Fag", "Område", "Frikode"]

    # Sikre at key-kolonner finnes (i tilfelle upstream endringer)
    for _kc in key_cols:
        if _kc not in vareoversikt.columns:
            vareoversikt[_kc] = ""
    if not kategorisert.empty:
        for _kc in key_cols:
            if _kc not in kategorisert.columns:
                kategorisert[_kc] = ""

    def _safe_join_key(vals):
        out = []
        for v in vals:
            if v is None:
                out.append("")
            else:
                s = str(v)
                out.append("" if s.lower() == "nan" else s)
        return "\n".join(out)

    # Bygg nøkler trygt
    vareoversikt["_KEY"] = vareoversikt[key_cols].apply(lambda r: _safe_join_key(r.values.tolist()), axis=1)
    if not kategorisert.empty:
        kategorisert["_KEY"] = kategorisert[key_cols].apply(lambda r: _safe_join_key(r.values.tolist()), axis=1)
        gul_keys = set(vareoversikt.loc[flag_vare_gul, "_KEY"].unique())
        kat_gul = kategorisert["_KEY"].isin(gul_keys)
    else:
        kat_gul = pd.Series([], dtype=bool)

    # Skriv til Excel
    YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        kategorisert.drop(columns=["_KEY"], errors="ignore").to_excel(writer, sheet_name="Kategorisert", index=False)
        vareoversikt.drop(columns=["_KEY"], errors="ignore").to_excel(writer, sheet_name="Vareoversikt", index=False)
        ukjente_varer.to_excel(writer, sheet_name="Ukjente varer", index=False)
        nobb_resultat.to_excel(writer, sheet_name="Varer med variasjon", index=False)

    # Legg på gult fyll
    wb = load_workbook(output_file)
    ws_vare = wb["Vareoversikt"]
    ws_kat = wb["Kategorisert"]
    ws_var = wb["Varer med variasjon"]

    n_cols_vare = ws_vare.max_column
    for i, gul in enumerate(flag_vare_gul, start=2):
        if gul:
            for c in range(1, n_cols_vare + 1):
                cell = ws_vare.cell(row=i, column=c)
                if cell.value not in (None, ""):
                    cell.fill = YELLOW

    n_cols_kat = ws_kat.max_column
    for i, gul in enumerate(kat_gul, start=2):
        if gul:
            for c in range(1, n_cols_kat + 1):
                cell = ws_kat.cell(row=i, column=c)
                if cell.value not in (None, ""):
                    cell.fill = YELLOW

    ws_vare.freeze_panes = "A2"
    ws_kat.freeze_panes = "A2"
    ws_var.freeze_panes = "A2"
    wb.save(output_file)


# ============================================================================
# DEL 4: STREAMLIT-APP
# ============================================================================

ProgressCallback = Optional[Callable[[float, str], None]]


def _safe_filename(name: str, fallback: str) -> str:
    """Gjør et opplastet filnavn trygt å bruke i en midlertidig mappe."""
    filename = Path(name or fallback).name
    filename = re.sub(r"[^A-Za-z0-9._-]+", "_", filename)
    return filename or fallback


def _write_uploaded_file(uploaded_file: Any, target: Path) -> None:
    """Skriv en Streamlit UploadedFile til disk i den midlertidige arbeidsmappen."""
    target.write_bytes(uploaded_file.getvalue())


def _uploaded_files_signature(pdf_files: List[Any], category_file: Any) -> str:
    """Signatur som hindrer at et gammelt resultat vises etter at inputfilene endres."""
    digest = hashlib.sha256()
    for uploaded_file in [*pdf_files, category_file]:
        digest.update((uploaded_file.name or "").encode("utf-8", errors="ignore"))
        digest.update(uploaded_file.getvalue())
    return digest.hexdigest()


def _count_output_rows(output_file: Path) -> Dict[str, int]:
    """Tell datarader i hvert resultatark, uten å ta med overskriftsraden."""
    workbook = load_workbook(output_file, read_only=True, data_only=True)
    try:
        return {
            sheet_name: max(workbook[sheet_name].max_row - 1, 0)
            for sheet_name in workbook.sheetnames
        }
    finally:
        workbook.close()


def process_uploaded_files(
    pdf_uploads: List[Any],
    category_upload: Any,
    debug: bool = False,
    progress_callback: ProgressCallback = None,
) -> Dict[str, Any]:
    """
    Kjør hele behandlingsflyten og returner Excel-filen som bytes.

    Alle opplastede filer lagres kun i en midlertidig mappe under kjøringen.
    Mappen slettes automatisk når behandlingen er ferdig.
    """
    if not pdf_uploads:
        raise ValueError("Last opp minst én PDF-faktura.")
    if category_upload is None:
        raise ValueError("Last opp kategoriseringsfilen.")

    log_buffer = io.StringIO()

    def report(value: float, message: str) -> None:
        if progress_callback:
            progress_callback(max(0.0, min(float(value), 1.0)), message)

    try:
        with redirect_stdout(log_buffer):
            with tempfile.TemporaryDirectory(prefix="optimera_streamlit_") as temp_dir:
                work_dir = Path(temp_dir)

                category_suffix = Path(category_upload.name or "").suffix.lower()
                if category_suffix not in {".xlsx", ".xlsm"}:
                    raise ValueError(
                        "Kategoriseringsfilen må være en .xlsx- eller .xlsm-fil."
                    )

                category_path = work_dir / (
                    "kategorisering" + category_suffix
                )
                _write_uploaded_file(category_upload, category_path)

                # Gi en tydelig feil før PDF-behandlingen dersom obligatorisk ark mangler.
                category_workbook = pd.ExcelFile(category_path, engine="openpyxl")
                if "Linjer" not in category_workbook.sheet_names:
                    available = ", ".join(category_workbook.sheet_names) or "ingen"
                    raise ValueError(
                        "Kategoriseringsfilen mangler arket 'Linjer'. "
                        f"Tilgjengelige ark: {available}."
                    )

                parser = InvoiceParser(debug=debug)
                all_parsed_data: List[Dict] = []
                file_results: List[Dict[str, Any]] = []
                total_relevant_pages = 0
                number_of_files = len(pdf_uploads)

                for index, pdf_upload in enumerate(pdf_uploads, start=1):
                    display_name = pdf_upload.name or f"faktura_{index}.pdf"
                    safe_name = _safe_filename(display_name, f"faktura_{index}.pdf")
                    if not safe_name.lower().endswith(".pdf"):
                        safe_name += ".pdf"
                    pdf_path = work_dir / f"{index:03d}_{safe_name}"
                    _write_uploaded_file(pdf_upload, pdf_path)

                    report(
                        0.05 + (index - 1) / max(number_of_files, 1) * 0.60,
                        f"Behandler {display_name} ({index} av {number_of_files}) …",
                    )

                    try:
                        filtered_pages = get_filtered_pages_from_pdf(str(pdf_path))
                        total_relevant_pages += len(filtered_pages)

                        if not filtered_pages:
                            file_results.append(
                                {
                                    "Fil": display_name,
                                    "Relevante sider": 0,
                                    "Varelinjer": 0,
                                    "Status": "Hoppet over – fant ikke 'Side X av Y'",
                                }
                            )
                            continue

                        parsed_rows = parser.parse_pdf_filtered(
                            str(pdf_path), filtered_pages
                        )
                        all_parsed_data.extend(parsed_rows)
                        file_results.append(
                            {
                                "Fil": display_name,
                                "Relevante sider": len(filtered_pages),
                                "Varelinjer": len(parsed_rows),
                                "Status": "OK",
                            }
                        )
                    except Exception as exc:
                        raise RuntimeError(
                            f"Kunne ikke behandle '{display_name}': {exc}"
                        ) from exc

                if not all_parsed_data:
                    raise ValueError(
                        "Ingen fakturalinjer ble funnet. Kontroller at PDF-ene er "
                        "Optimera-fakturaer med teksten 'Side X av Y'."
                    )

                report(0.70, "Forbereder varelinjene …")
                df_parsed = pd.DataFrame(all_parsed_data)

                # Behold samme normalisering som i skrivebordsversjonen.
                for column in df_parsed.columns:
                    df_parsed[column] = (
                        df_parsed[column]
                        .astype(str)
                        .replace("nan", "")
                        .fillna("")
                    )
                df_parsed = df_parsed.rename(
                    columns={"Nobb-nr": "Nobbnr", "Art-nr": "Artnr"}
                )

                invoice_column = finn_kolonnenavn(df_parsed, "fakturanr")
                invoice_count = (
                    int(df_parsed[invoice_column].replace("", pd.NA).nunique())
                    if invoice_column
                    else 0
                )

                output_path = work_dir / "Optimera_fakturaresultat.xlsx"
                report(0.78, "Kategoriserer varelinjer og bygger Excel-filen …")
                kategoriser_og_eksporter(
                    df_parsed,
                    str(category_path),
                    str(output_path),
                    debug=debug,
                )

                report(0.94, "Leser resultat og lager forhåndsvisning …")
                row_counts = _count_output_rows(output_path)
                preview = pd.read_excel(
                    output_path,
                    sheet_name="Vareoversikt",
                    engine="openpyxl",
                    nrows=100,
                )
                excel_bytes = output_path.read_bytes()

                report(1.0, "Ferdig")
                return {
                    "excel_bytes": excel_bytes,
                    "parsed_lines": len(all_parsed_data),
                    "invoice_count": invoice_count,
                    "pdf_count": number_of_files,
                    "relevant_pages": total_relevant_pages,
                    "row_counts": row_counts,
                    "file_results": file_results,
                    "preview": preview,
                    "log": log_buffer.getvalue(),
                }
    except Exception as exc:
        # Loggen vises bare i appen når brukeren har slått på feilsøkingsmodus.
        setattr(exc, "debug_log", log_buffer.getvalue())
        raise


def _normalise_output_name(name: str) -> str:
    filename = _safe_filename(name.strip(), "Optimera_fakturaresultat.xlsx")
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    return filename


def main() -> None:
    st.set_page_config(
        page_title="Optimera fakturabehandling",
        page_icon="📄",
        layout="wide",
    )

    st.title("Optimera fakturabehandling")
    st.write(
        "Last opp én eller flere Optimera-fakturaer og kategoriseringsfilen. "
        "Appen parser fakturalinjene, kobler dem mot kategoriene og lager én "
        "Excel-fil som kan lastes ned."
    )

    with st.sidebar:
        st.header("Innstillinger")
        debug = st.checkbox(
            "Feilsøkingsmodus",
            value=False,
            help="Vis teknisk logg og full feilmelding. Brukes ved feilsøking.",
        )
        output_name = st.text_input(
            "Navn på resultatfil",
            value="Optimera_fakturaresultat.xlsx",
        )
        st.divider()
        st.caption(
            "Filene behandles i en midlertidig mappe og slettes fra "
            "arbeidsområdet når kjøringen er ferdig."
        )

    left, right = st.columns(2)
    with left:
        st.subheader("1. Fakturaer")
        pdf_uploads = st.file_uploader(
            "Last opp PDF-fakturaer",
            type=["pdf"],
            accept_multiple_files=True,
            help="Du kan velge flere PDF-filer samtidig.",
        )

    with right:
        st.subheader("2. Kategoriseringsfil")
        category_upload = st.file_uploader(
            "Last opp Excel-filen med kategorier",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            help="Filen må inneholde arket 'Linjer'.",
        )

    if pdf_uploads:
        total_mb = sum(upload.size for upload in pdf_uploads) / (1024 * 1024)
        st.info(
            f"Valgt {len(pdf_uploads)} PDF-fil(er), totalt {total_mb:.1f} MB."
        )
        with st.expander("Vis valgte PDF-filer"):
            for uploaded_file in pdf_uploads:
                st.write(f"• {uploaded_file.name}")

    ready = bool(pdf_uploads) and category_upload is not None
    current_signature = (
        _uploaded_files_signature(pdf_uploads, category_upload) if ready else None
    )

    # Fjern et tidligere resultat dersom brukeren bytter inputfiler.
    if (
        "optimera_result" in st.session_state
        and st.session_state.get("optimera_signature") != current_signature
    ):
        st.session_state.pop("optimera_result", None)
        st.session_state.pop("optimera_signature", None)

    st.subheader("3. Behandle")
    run_clicked = st.button(
        "Behandle fakturaer",
        type="primary",
        disabled=not ready,
        use_container_width=True,
    )

    if not ready:
        st.caption("Last opp både PDF-fakturaer og kategoriseringsfil for å fortsette.")

    if run_clicked:
        progress = st.progress(0, text="Starter …")

        def update_progress(value: float, message: str) -> None:
            progress.progress(value, text=message)

        try:
            result = process_uploaded_files(
                pdf_uploads=pdf_uploads,
                category_upload=category_upload,
                debug=debug,
                progress_callback=update_progress,
            )
            st.session_state["optimera_result"] = result
            st.session_state["optimera_signature"] = current_signature
        except Exception as exc:
            st.session_state.pop("optimera_result", None)
            st.session_state.pop("optimera_signature", None)
            progress.empty()
            st.error(f"Behandlingen stoppet: {exc}")
            if debug:
                debug_log = getattr(exc, "debug_log", "")
                if debug_log:
                    with st.expander("Teknisk logg", expanded=True):
                        st.code(debug_log)
                with st.expander("Full feilmelding", expanded=True):
                    st.code("".join(traceback.format_exception(exc)))
        else:
            progress.progress(1.0, text="Ferdig")

    result = st.session_state.get("optimera_result")
    if result:
        st.success("Fakturabehandlingen er ferdig.")

        row_counts = result["row_counts"]
        metric_columns = st.columns(5)
        metric_columns[0].metric("PDF-filer", result["pdf_count"])
        metric_columns[1].metric("Fakturaer", result["invoice_count"])
        metric_columns[2].metric("Varelinjer", result["parsed_lines"])
        metric_columns[3].metric(
            "Ukjente varer", row_counts.get("Ukjente varer", 0)
        )
        metric_columns[4].metric(
            "Variasjonsrader", row_counts.get("Varer med variasjon", 0)
        )

        st.download_button(
            label="Last ned Excel-resultatet",
            data=result["excel_bytes"],
            file_name=_normalise_output_name(output_name),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

        with st.expander("Resultat per PDF"):
            st.dataframe(
                pd.DataFrame(result["file_results"]),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Forhåndsvis Vareoversikt (maks. 100 rader)"):
            st.dataframe(
                result["preview"],
                use_container_width=True,
                hide_index=True,
            )

        st.caption(
            "Excel-filen inneholder arkene Kategorisert, Vareoversikt, "
            "Ukjente varer og Varer med variasjon. Gul markering betyr at "
            "varen har variasjon i kategoriseringen."
        )

        if debug and result.get("log"):
            with st.expander("Teknisk logg"):
                st.code(result["log"])


if __name__ == "__main__":
    main()
