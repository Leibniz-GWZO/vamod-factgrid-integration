import re
import string
from pathlib import Path
import pandas as pd


def _split_on_dash_segments(value):
    """Splits on dashes allowing arbitrary spaces around them."""
    return re.split(r"\s*-\s*", value)


def _join_date_parts(parts):
    """Join parts from index 2 onward into a date with no spaces around dashes."""
    return "-".join(p.strip() for p in parts[2:])


def clean_Lde(value):
    if not isinstance(value, str):
        return value
    parts = _split_on_dash_segments(value)
    if len(parts) < 3:
        return value
    first = parts[0].strip()
    second = parts[1].strip()
    date = _join_date_parts(parts)
    return f"{first} - {second} - {date}"


# characters to trim as punctuation (ASCII punctuation plus some common Unicode punctuation)
_EXTRA_PUNCT = '«»—…“”„'
_PUNCT_PATTERN = re.escape(string.punctuation + _EXTRA_PUNCT)


def _strip_edge_punctuation(s):
    # remove punctuation at the beginning and end of the string
    s = re.sub(rf'^[{_PUNCT_PATTERN}\s]+', '', s)
    s = re.sub(rf'[{_PUNCT_PATTERN}\s]+$', '', s)
    return s


def clean_Lukr(value):
    if not isinstance(value, str):
        return value
    parts = _split_on_dash_segments(value)
    if len(parts) < 3:
        return value
    first = parts[0].strip()
    second = _strip_edge_punctuation(parts[1])
    date = _join_date_parts(parts)
    return f"{first} - {second} - {date}"


# month name to number mapping (German, common abbreviations)
_MONTH_MAP = {
    'januar': 1, 'jan': 1,
    'februar': 2, 'feb': 2,
    'märz': 3, 'maerz': 3, 'mrz': 3, 'mar': 3, 'mär': 3,
    'april': 4, 'apr': 4,
    'mai': 5,
    'juni': 6, 'jun': 6,
    'juli': 7, 'jul': 7,
    'august': 8, 'aug': 8,
    'september': 9, 'sep': 9, 'sept': 9,
    'oktober': 10, 'okt': 10,
    'november': 11, 'nov': 11,
    'dezember': 12, 'dez': 12, 'dec': 12
}


def _normalize_token(tok: str) -> str:
    t = tok.lower().strip().strip('.')
    # normalize common umlauts to plain forms that might appear in sources
    t = t.replace('ä', 'a').replace('ö', 'o').replace('ü', 'u').replace('ß', 'ss')
    return t


def parse_date_field(value):
    """Parse a messy German date string into +YYYY-MM-DDT00:00:00Z/accuracy.

    Accuracy codes used:
    /11 = day
    /10 = month
    /9  = year
    /8  = decade (when year not a 4-digit number; derive from first two digits)
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return s

    # Handle specific cases with exact string matching FIRST (before normalization)
    if s == "Nach Apr. 1374, vor Ende 1375" or s == "+1370-00-00T00:00:00Z/8":  # checked
        return "+1370-00-00T00:00:00Z/8"
    if s == "Nov. 12" or s == "+1419-11-12T00:00:00Z/11":  # checked
        return "+1419-11-12T00:00:00Z/11"
    if s == "1423 Juli [13 bis 27]" or s == "+1423-07-00T00:00:00Z/10":  # checked
        return "+1423-07-00T00:00:00Z/10"
    if s == "1418–1425" or s == "+1410-00-00T00:00:00Z/8":  # checked
        return "+1410-00-00T00:00:00Z/8"
    if s == "(1424–1427)" or s == "+1420-00-00T00:00:00Z/8":  # checked
        return "+1420-00-00T00:00:00Z/8"
    if s == "(Ende März bis Anfang Apr.) 1423–1427" or s == "+1420-00-00T00:00:00Z/8":  # checked
        return "+1420-00-00T00:00:00Z/8"
    if s == "Dez. 10" or s == "+1431-12-10T00:00:00Z/11":  # checked
        return "+1431-12-10T00:00:00Z/11"
    if s == "Feb. 15" or s == "+1432-02-15T00:00:00Z/11":  # checked
        return "+1432-02-15T00:00:00Z/11"

    # normalize quote characters and strip leading apostrophes/double quotes
    s_norm = s.replace('\u2019', "'").replace('\u2018', "'").replace('`', "'").replace('´', "'")
    s_norm = s_norm.replace('\u201c', '"').replace('\u201d', '"')
    s_norm = s_norm.lstrip('"\'')

    # if already in the target format (starts with '+'), leave unchanged
    if s_norm.startswith('+'):
        return value

    # try to find a 4-digit year
    y_match = re.search(r'(\d{4})', s_norm)
    if y_match:
        year = int(y_match.group(1))
        year_str = f"{year:04d}"
    else:
        # fallback: try to find two leading digits (decade-level information)
        two = re.search(r'\b(\d{2})\b', s_norm)
        if two:
            yy = int(two.group(1))
            year_str = f"{yy:02d}00"
            return f"+{year_str}-00-00T00:00:00Z/8"
        return s

    lowered = s_norm.lower()
    # create a normalized version (remove dots, normalize umlauts) for reliable searching
    norm = _normalize_token(lowered)

    # detect all months mentioned in the normalized string
    found_months = []
    month_positions = []
    for key, num in _MONTH_MAP.items():
        key_norm = _normalize_token(key)
        # use word boundaries on the normalized string
        m = re.search(r"\b" + re.escape(key_norm) + r"\b", norm)
        if m:
            if num not in found_months:
                found_months.append(num)
                month_positions.append((m.start(), m.end(), num))

    # check for uncertainty markers like (?) that reduce precision
    has_uncertainty = '(?)' in s_norm or '?' in s_norm
    
    # if multiple different months found or explicit slash/range between month names -> month ambiguous
    if len(found_months) > 1 or ('/' in norm and any(k in norm for k in ['jan', 'feb', 'mar', 'apr', 'mai', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dez'])):
        # fall back to year-only
        return f"+{year_str}-00-00T00:00:00Z/9"

    # if exactly one month recognized
    if len(found_months) == 1:
        month = found_months[0]
        # if uncertainty marker present, reduce to month precision
        if has_uncertainty:
            return f"+{year_str}-{month:02d}-00T00:00:00Z/10"
            
        # find position of the month in normalized string to look for day tokens after it
        pos = None
        for a, b, num in month_positions:
            if num == month:
                pos = b
                break
        if pos is None:
            # cannot determine position; still return month accuracy
            return f"+{year_str}-{month:02d}-01T00:00:00Z/10"

        # substring after month in normalized string to find day or ranges
        sub = norm[pos:]
        # day range like 26-31 or 26–31 -> ambiguous, use month accuracy
        range_match = re.search(r"(\d{1,2})\s*[–\-]\s*(\d{1,2})", sub)
        if range_match:
            return f"+{year_str}-{month:02d}-00T00:00:00Z/10"

        # single day
        day_match = re.search(r"(\d{1,2})", sub)
        if day_match:
            day = int(day_match.group(1))
            return f"+{year_str}-{month:02d}-{day:02d}T00:00:00Z/11"

        # no day found -> month accuracy
        return f"+{year_str}-{month:02d}-00T00:00:00Z/10"

    # no month found -> return year accuracy
    return f"+{year_str}-00-00T00:00:00Z/9"


def format_parsed_date(parsed):
    """Format the parsed date string to YYYY-MM-DD, YYYY-MM, YYYY, or YYYY-YYYY for decades based on accuracy."""
    if not parsed or not isinstance(parsed, str) or not parsed.startswith('+'):
        return parsed
    match = re.match(r'\+(\d{4})-(\d{2})-(\d{2})T.*Z/(\d+)', parsed)
    if match:
        year, month, day, acc = match.groups()
        if acc == '11':
            return f"{year}-{month}-{day}"
        elif acc == '10':
            return f"{year}-{month}"
        elif acc == '9':
            return year
        elif acc == '8':
            return f"{year}-{int(year) + 9}"  # Decade as YYYY-YYYY
    return parsed


def clean_Len(value):
    if not isinstance(value, str):
        return value
    parts = _split_on_dash_segments(value)
    if len(parts) < 3:
        return value
    first = parts[0].strip()
    second = parts[1].strip()
    date = _join_date_parts(parts)
    return f"{first} - {second} - {date}"


def clean_Lpl(value):
    if not isinstance(value, str):
        return value
    parts = _split_on_dash_segments(value)
    if len(parts) < 3:
        return value
    first = parts[0].strip()
    second = parts[1].strip()
    date = _join_date_parts(parts)
    return f"{first} - {second} - {date}"


def clean_date_column(df: pd.DataFrame, src_col: str, dst_col: str = 'Datum_v9') -> pd.DataFrame:
    """Add a new column dst_col to df containing parsed date strings.

    The function preserves existing dataframe and returns it (in-place modification of df).
    """
    if src_col not in df.columns:
        return df
    df[dst_col] = df[src_col].apply(parse_date_field)
    return df


def main():
    base = Path(__file__).resolve().parent
    file_path = base.joinpath('..', 'datasheets', 'Imports', 'Urkunden_Metadatenliste_v8.xlsx').resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    df = pd.read_excel(file_path)

    # parse German date column into Datum_v9
    date_col = 'Datum (P106) (P43 Datum vor, P41 Datum nach)'
    df = clean_date_column(df, date_col, 'Datum_v9')

    # clean Lde, Len, Lpl, Lukr columns using formatted date from Datum_v9
    for index, row in df.iterrows():
        formatted_date = format_parsed_date(row['Datum_v9'])
        for col in ['Lde', 'Len', 'Lpl', 'Lukr']:
            if col in df.columns:
                value = row[col]
                if isinstance(value, str):
                    parts = _split_on_dash_segments(value)
                    if len(parts) >= 3:
                        first = parts[0].strip()
                        if col == 'Lukr':
                            second = _strip_edge_punctuation(parts[1])
                        else:
                            second = parts[1].strip()
                        df.at[index, col] = f"{first} - {second} - {formatted_date}"

    # save original cleaned file and also write a new v9 Excel file
    df.to_excel(file_path, index=False)
    v9_path = base.joinpath('..', 'datasheets', 'Imports', 'Urkunden_Metadatenliste_v9.xlsx').resolve()
    df.to_excel(v9_path, index=False)


if __name__ == '__main__':
    main()