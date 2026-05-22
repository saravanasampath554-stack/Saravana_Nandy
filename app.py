"""
Measurement Verifier Backend
-----------------------------
Accepts PDF/image uploads containing tables with columns: No, L, B, D, Content/Area.
Extracts table data, verifies No×L×B×D == Content/Area, and returns results.
"""

import os
import re
import tempfile
import traceback
import uuid
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd

import hashlib
import json

FRONTEND_DIR = os.path.dirname(__file__)
SAVED_BILLS_DIR = os.path.join(os.path.dirname(__file__), 'saved_bills')
ACCOUNTS_FILE = os.path.join(os.path.dirname(__file__), 'accounts.json')

app = Flask(__name__, static_folder=FRONTEND_DIR)
CORS(app)

UPLOAD_FOLDER = tempfile.mkdtemp()
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'bmp', 'tiff', 'xlsx', 'xls', 'csv'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_number(val):
    """Safely parse a numeric value from extracted text.
    Supports multiplication expressions like '2*3', '2x3x4', '2×3'.
    """
    if val is None:
        return None
    # Handle pandas Series (from duplicate column names)
    if isinstance(val, pd.Series):
        val = val.iloc[0]
    try:
        if hasattr(val, '__len__') and not isinstance(val, str):
            return None
        if pd.isna(val):
            return None
    except Exception:
        return None
    val = str(val).strip().replace(',', '')

    # Check if it's a multiplication expression (e.g. 2*3, 2x3x4, 2×3)
    if re.search(r'[*xX×]', val):
        parts = re.split(r'[*xX×]', val)
        result = 1.0
        for part in parts:
            part = part.strip()
            part = re.sub(r'[^\d.\-]', '', part)
            if part == '' or part == '-':
                return None
            try:
                result *= float(part)
            except ValueError:
                return None
        return result

    # Simple number
    val = re.sub(r'[^\d.\-]', '', val)
    if val == '' or val == '-':
        return None
    try:
        return float(val)
    except ValueError:
        return None


def normalize_column_name(col):
    """Map extracted column names to standard names."""
    col_clean = str(col).strip().upper()
    col_clean = re.sub(r'[^A-Z0-9]', '', col_clean)

    mapping = {
        'NO': 'No', 'NOS': 'No', 'NUMBER': 'No', 'NUM': 'No',
        'SR': 'No', 'SRNO': 'No', 'SN': 'No', 'SLNO': 'No', 'SNO': 'No',
        'SERIAL': 'No', 'SERIALNO': 'No', 'SERNO': 'No',
        'L': 'L', 'LENGTH': 'L', 'LEN': 'L',
        'B': 'B', 'BREADTH': 'B', 'WIDTH': 'B', 'W': 'B', 'BRI': 'B',
        'D': 'D', 'DEPTH': 'D', 'HEIGHT': 'D', 'H': 'D', 'HT': 'D',
        'CONTENT': 'Content', 'AREA': 'Content', 'QUANTITY': 'Content',
        'CONTENTS': 'Content', 'AMOUNT': 'Content', 'QTY': 'Content',
        'CONT': 'Content', 'CONTENTAREA': 'Content', 'VOL': 'Content',
        'VOLUME': 'Content', 'RESULT': 'Content', 'TOTAL': 'Content',
    }
    return mapping.get(col_clean, col)


def try_positional_columns(df):
    """If column names can't be matched, try assigning by position.
    Common measurement sheet layouts:
      - 5 cols: No, L, B, D, Content
      - 6 cols: Sr/Particulars, No, L, B, D, Content
      - 7 cols: Sr, Particulars, No, L, B, D, Content
      - 8 cols: Sr, Particulars, No, L, B, D, Content, Rate/Amount
    """
    ncols = len(df.columns)
    if ncols < 5:
        return df

    # Check how many columns already matched
    required = ['No', 'L', 'B', 'D', 'Content']
    matched = sum(1 for c in df.columns if c in required)
    if matched >= 3:
        return df  # Already matched enough

    # Try to detect numeric columns from data
    numeric_cols = []
    sample_size = min(10, len(df))
    for i, col in enumerate(df.columns):
        numeric_count = 0
        for val in df[col].head(sample_size):
            if parse_number(val) is not None:
                numeric_count += 1
        if sample_size > 0 and numeric_count > sample_size * 0.3:
            numeric_cols.append(i)

    # Standard positional mappings based on column count
    col_maps = {
        5: {0: 'No', 1: 'L', 2: 'B', 3: 'D', 4: 'Content'},
        6: {1: 'No', 2: 'L', 3: 'B', 4: 'D', 5: 'Content'},
        7: {2: 'No', 3: 'L', 4: 'B', 5: 'D', 6: 'Content'},
        8: {2: 'No', 3: 'L', 4: 'B', 5: 'D', 6: 'Content'},
    }

    if ncols in col_maps:
        new_cols = list(df.columns)
        for pos, name in col_maps[ncols].items():
            if pos < len(new_cols):
                new_cols[pos] = name
        df.columns = new_cols
    elif len(numeric_cols) >= 5:
        # Use the last 5 numeric columns as No, L, B, D, Content
        positions = numeric_cols[-5:]
        names = ['No', 'L', 'B', 'D', 'Content']
        new_cols = list(df.columns)
        for pos, name in zip(positions, names):
            new_cols[pos] = name
        df.columns = new_cols

    return df


def extract_tables_from_pdf(filepath):
    """Extract tables from PDF using tabula-py."""
    import tabula
    tables = tabula.read_pdf(filepath, pages='all', multiple_tables=True, lattice=True)
    if not tables:
        tables = tabula.read_pdf(filepath, pages='all', multiple_tables=True, stream=True)
    return tables


def extract_tables_from_excel(filepath):
    """Extract tables from Excel files."""
    xls = pd.ExcelFile(filepath)
    tables = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet)
        # Drop fully empty rows/columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        if not df.empty:
            tables.append(df)
    return tables


def preprocess_image(filepath):
    """Preprocess image with OpenCV for better OCR on handwritten content."""
    import cv2
    img = cv2.imread(filepath)
    if img is None:
        return filepath

    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Resize if too small (EasyOCR works better on larger images)
    h, w = gray.shape
    if max(h, w) < 2000:
        scale = 2000 / max(h, w)
        gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    # Denoise
    gray = cv2.fastNlMeansDenoising(gray, h=15)

    # Save preprocessed grayscale (no binarization - EasyOCR works better on grayscale)
    processed_path = filepath + '_processed.png'
    cv2.imwrite(processed_path, gray)
    return processed_path


# Initialize EasyOCR reader once (lazy, downloads model on first use)
_easyocr_reader = None

def _get_reader():
    global _easyocr_reader
    if _easyocr_reader is None:
        import easyocr
        _easyocr_reader = easyocr.Reader(['en'], gpu=False)
    return _easyocr_reader


def _clean_ocr_number(text):
    """Fix common OCR misreads in handwritten numbers.
    Only applies corrections if the text already looks mostly numeric.
    Returns cleaned string or None if the text doesn't look like a number.
    """
    if not text:
        return None
    # Remove spaces
    t = text.strip().replace(' ', '')

    if not t:
        return None

    # Count how many characters are already digits or decimal/sign
    digit_like = sum(1 for ch in t if ch in '0123456789.,-*')
    total_chars = len(t)

    # If less than 40% of characters are digit-like, it's probably not a number
    if total_chars > 2 and digit_like / total_chars < 0.4:
        return None

    # If text is mostly letters with no digits at all, reject
    if sum(1 for ch in t if ch in '0123456789') == 0 and total_chars > 1:
        return None

    # Common OCR character substitutions for digits
    replacements = {
        'S': '5', 's': '5',
        'I': '1', 'l': '1', 'i': '1',
        'O': '0', 'o': '0',
        '+': '7',
        'G': '6',
        'Z': '2', 'z': '2',
        'q': '9',
        '(': '1', ')': '1',
        '{': '1', '}': '1',
        '#': '',
        '&': '',
    }

    result = []
    for idx, ch in enumerate(t):
        if ch in '0123456789.,':
            result.append(ch)
        elif ch == '-' and idx == 0:
            result.append('-')  # Negative sign
        elif ch == '-':
            result.append('.')  # Dash in middle → decimal point
        elif ch == '*' and idx > 0:
            result.append('.')  # Asterisk between digits → decimal point
        elif ch in replacements:
            result.append(replacements[ch])
        # skip other non-digit chars

    cleaned = ''.join(result)

    if not cleaned or cleaned in ('.', '-', ',', '..'):
        return None

    # Remove commas
    cleaned = cleaned.replace(',', '')

    # Handle multiple dots - keep only the first
    if cleaned.count('.') > 1:
        parts = cleaned.split('.')
        cleaned = parts[0] + '.' + ''.join(parts[1:])

    return cleaned


def extract_tables_from_image(filepath):
    """Extract tables from image using EasyOCR + OpenCV with column-aware parsing."""
    import cv2

    # Preprocess the image
    processed_path = preprocess_image(filepath)

    try:
        reader = _get_reader()
        results = reader.readtext(processed_path, paragraph=False, width_ths=0.5)

        if not results:
            return [], "No text detected in image"

        # Build entries with positions
        entries = []
        for bbox, text, conf in results:
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            x_center = (bbox[0][0] + bbox[2][0]) / 2
            entries.append({'text': text.strip(), 'x': x_center, 'y': y_center, 'conf': conf})

        if not entries:
            return [], "No text detected in image"

        img = cv2.imread(processed_path)
        img_height = img.shape[0] if img is not None else 1000
        img_width = img.shape[1] if img is not None else 2000
        row_threshold = img_height * 0.03  # 3% of image height for row grouping

        # --- Step 1: Find column headers to establish column x-positions ---
        header_keywords = {
            'No': ['no', 'no.', 'nos', 'nos.'],
            'B': ['b', 'breadth'],
            'D': ['d', 'depth'],
            'Content': ['content', 'contents', 'area', 'areas'],
        }
        # Also detect non-numeric columns to exclude them
        skip_keywords = ['particulars', 'partiulars', 'particular', 'rate', 'amount', 'remark', 'remarks']

        col_x = {}  # column name -> x position
        skip_x = []  # x positions of non-numeric columns (Particulars, Rate, etc.)
        header_y = None

        for e in entries:
            txt = e['text'].lower().strip().rstrip('.')
            # Check for skip columns
            for kw in skip_keywords:
                if kw in txt.lower():
                    skip_x.append(e['x'])
                    if header_y is None:
                        header_y = e['y']
                    break
            # Check for data columns
            for col, kws in header_keywords.items():
                if txt in kws or txt.rstrip('.') in kws:
                    col_x[col] = e['x']
                    if header_y is None:
                        header_y = e['y']
                    else:
                        header_y = min(header_y, e['y'])

        # Estimate L column position: between No and B
        if 'No' in col_x and 'B' in col_x:
            col_x['L'] = (col_x['No'] + col_x['B']) / 2
        elif 'No' in col_x:
            col_x['L'] = col_x['No'] + (img_width * 0.12)

        # If we couldn't detect headers, try to infer from data clusters
        if len(col_x) < 3:
            # Collect x positions of entries that look numeric
            numeric_xs = []
            for e in entries:
                cleaned = _clean_ocr_number(e['text'])
                try:
                    float(cleaned)
                    numeric_xs.append(e['x'])
                except (ValueError, TypeError):
                    pass

            if numeric_xs:
                # Cluster x positions to find columns
                numeric_xs.sort()
                clusters = []
                current_cluster = [numeric_xs[0]]
                for x in numeric_xs[1:]:
                    if x - current_cluster[-1] < img_width * 0.08:
                        current_cluster.append(x)
                    else:
                        clusters.append(sum(current_cluster) / len(current_cluster))
                        current_cluster = [x]
                clusters.append(sum(current_cluster) / len(current_cluster))

                # Assign the last 5 (or fewer) clusters as No, L, B, D, Content
                col_names = ['No', 'L', 'B', 'D', 'Content']
                if len(clusters) >= 5:
                    for i, name in enumerate(col_names):
                        col_x[name] = clusters[-(5 - i)]
                elif len(clusters) >= 4:
                    for i, name in enumerate(['L', 'B', 'D', 'Content']):
                        col_x[name] = clusters[-(4 - i)]

        if not col_x:
            return [], "Could not detect table columns in image"

        # Ensure we have all 5 columns — fill gaps by interpolation
        all_cols = ['No', 'L', 'B', 'D', 'Content']
        known_positions = {c: col_x[c] for c in all_cols if c in col_x}
        if len(known_positions) >= 2:
            sorted_known = sorted(known_positions.items(), key=lambda x: x[1])
            # Estimate spacing
            spacings = []
            for i in range(len(sorted_known) - 1):
                idx1 = all_cols.index(sorted_known[i][0])
                idx2 = all_cols.index(sorted_known[i + 1][0])
                spacings.append((sorted_known[i + 1][1] - sorted_known[i][1]) / (idx2 - idx1))
            avg_spacing = sum(spacings) / len(spacings)

            for col in all_cols:
                if col not in col_x:
                    # Find nearest known column and extrapolate
                    for known_col, known_pos in known_positions.items():
                        idx_diff = all_cols.index(col) - all_cols.index(known_col)
                        col_x[col] = known_pos + idx_diff * avg_spacing
                        break

        # --- Step 2: Define column boundaries ---
        col_centers = {c: col_x[c] for c in all_cols if c in col_x}
        sorted_cols = sorted(col_centers.items(), key=lambda x: x[1])

        # Find the rightmost skip column (Particulars) to use as left cutoff
        particulars_right = 0
        if skip_x and 'No' in col_x:
            # Particulars right edge = midpoint between Particulars center and No center
            max_skip = max(skip_x)
            if max_skip < col_x['No']:
                particulars_right = (max_skip + col_x['No']) / 2

        col_bounds = {}  # column name -> (left_bound, right_bound)
        for i, (col, cx) in enumerate(sorted_cols):
            if i == 0:
                left = max(particulars_right, 0)  # Don't go into Particulars zone
            else:
                left = (sorted_cols[i - 1][1] + cx) / 2
            if i == len(sorted_cols) - 1:
                right = img_width
            else:
                right = (cx + sorted_cols[i + 1][1]) / 2
            col_bounds[col] = (left, right)

        # --- Step 3: Filter entries below header, assign to columns ---
        if header_y is None:
            # Guess header is in top 35%
            header_y = img_height * 0.25

        data_entries = [e for e in entries if e['y'] > header_y + row_threshold]

        # Assign each entry to its column
        for e in data_entries:
            assigned = None
            for col, (left, right) in col_bounds.items():
                if left <= e['x'] < right:
                    assigned = col
                    break
            e['column'] = assigned

        # --- Step 4: Group entries into rows ---
        data_entries.sort(key=lambda e: e['y'])
        rows = []
        current_row = [data_entries[0]] if data_entries else []

        for entry in data_entries[1:]:
            # Check if this entry is on the same row (close y)
            avg_y = sum(e['y'] for e in current_row) / len(current_row)
            if abs(entry['y'] - avg_y) < row_threshold:
                current_row.append(entry)
            else:
                rows.append(current_row)
                current_row = [entry]
        if current_row:
            rows.append(current_row)

        # --- Step 5: Merge fragments and build table ---
        table_data = []
        for row_entries in rows:
            # Group entries by column, merge text for same column
            col_texts = {}
            for e in row_entries:
                col = e.get('column')
                if col and col in all_cols:
                    if col not in col_texts:
                        col_texts[col] = []
                    col_texts[col].append(e)

            if not col_texts:
                continue

            # Merge fragments in same column (sort by x, concatenate)
            row_data = {}
            for col in all_cols:
                if col in col_texts:
                    fragments = sorted(col_texts[col], key=lambda e: e['x'])

                    # Smart merging: if fragments are spatially separated and look like
                    # parts of a decimal (e.g. "0" + "350" → "0.350"), insert decimal point
                    if len(fragments) > 1:
                        parts = []
                        for fi, frag in enumerate(fragments):
                            parts.append(frag['text'])
                            if fi < len(fragments) - 1:
                                gap = fragments[fi + 1]['x'] - frag['x']
                                # If there's a notable gap and current part is short (1-2 chars)
                                # and looks like the integer part of a decimal
                                curr = frag['text'].strip()
                                if gap > img_width * 0.02 and len(curr) <= 2:
                                    parts.append('.')
                        merged_text = ''.join(parts)
                    else:
                        merged_text = fragments[0]['text']

                    # Clean OCR misreads and extract number
                    cleaned = _clean_ocr_number(merged_text)
                    if cleaned is not None:
                        try:
                            row_data[col] = str(float(cleaned))
                        except (ValueError, TypeError):
                            row_data[col] = None
                    else:
                        # Check for multiplication expression (e.g. "1x1", "Ixl", "Lxi", "(Xi")
                        # Normalize common OCR misreads of digits
                        norm = merged_text
                        for old, new in [('I', '1'), ('l', '1'), ('i', '1'), ('L', '1'),
                                         ('(', '1'), ('t', '1'), ('|', '1'),
                                         ('X', 'x'), ('×', 'x')]:
                            norm = norm.replace(old, new)
                        if re.match(r'^[\d]+x[\d]+$', norm):
                            row_data[col] = norm  # Keep as expression like "1x1"
                        else:
                            row_data[col] = None
                else:
                    row_data[col] = None

            # Only include rows that have at least 2 numeric values
            num_count = sum(1 for v in row_data.values() if v is not None)
            if num_count >= 2:
                table_data.append(row_data)

        if not table_data:
            return [], "No numeric table data detected in image"

        df = pd.DataFrame(table_data, columns=all_cols)
        return [df], None

    finally:
        # Clean up preprocessed image
        if processed_path != filepath and os.path.exists(processed_path):
            os.remove(processed_path)



def process_table(df, sheet_name="Sheet"):
    """Process a single table: normalize columns, verify calculations."""
    # Normalize column names
    df.columns = [normalize_column_name(c) for c in df.columns]

    required = ['No', 'L', 'B', 'D', 'Content']
    found = [c for c in required if c in df.columns]
    missing = [c for c in required if c not in df.columns]

    # If not enough columns matched by name, try positional mapping
    if len(found) < 3:
        df = try_positional_columns(df)
        found = [c for c in required if c in df.columns]
        missing = [c for c in required if c not in df.columns]

    if len(found) < 3:
        return {
            'sheet_name': sheet_name,
            'error': f"Could not find enough columns. Found: {list(df.columns)}. Missing: {missing}",
            'rows': [],
            'summary': {}
        }

    results = []
    content_sum = 0.0
    correct_count = 0
    incorrect_count = 0

    # Deduplicate column names to avoid Series returns from row.get()
    seen = {}
    new_cols = []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df.columns = new_cols

    for idx, row in df.iterrows():
        no_val = parse_number(row.get('No'))
        l_val = parse_number(row.get('L'))
        b_val = parse_number(row.get('B'))
        d_val = parse_number(row.get('D'))
        content_val = parse_number(row.get('Content'))

        # Skip rows where all values are None (empty rows / headers)
        if all(v is None for v in [no_val, l_val, b_val, d_val, content_val]):
            continue

        # Calculate expected = No × L × B × D (treat None as 1)
        factors = [v for v in [no_val, l_val, b_val, d_val] if v is not None]
        if factors:
            expected = 1.0
            for f in factors:
                expected *= f
            expected = round(expected, 4)
        else:
            expected = None

        # Verify
        is_correct = None
        difference = None
        if expected is not None and content_val is not None:
            difference = round(content_val - expected, 4)
            is_correct = abs(difference) < 0.01  # tolerance
            if is_correct:
                correct_count += 1
            else:
                incorrect_count += 1

        if content_val is not None:
            content_sum += content_val

        results.append({
            'row': idx + 1,
            'No': no_val,
            'L': l_val,
            'B': b_val,
            'D': d_val,
            'content_given': content_val,
            'content_calculated': expected,
            'difference': difference,
            'is_correct': is_correct
        })

    return {
        'sheet_name': sheet_name,
        'error': None,
        'columns_found': found,
        'columns_missing': missing,
        'rows': results,
        'summary': {
            'total_rows': len(results),
            'correct': correct_count,
            'incorrect': incorrect_count,
            'skipped': len(results) - correct_count - incorrect_count,
            'content_sum': round(content_sum, 4)
        }
    }


@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': f'File type not allowed. Use: {", ".join(ALLOWED_EXTENSIONS)}'}), 400

    # Save with unique name to avoid collisions
    ext = file.filename.rsplit('.', 1)[1].lower()
    unique_name = f"{uuid.uuid4().hex}.{ext}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
    file.save(filepath)

    try:
        is_image = ext in ('png', 'jpg', 'jpeg', 'bmp', 'tiff')

        if ext in ('xlsx', 'xls'):
            tables = extract_tables_from_excel(filepath)
        elif ext == 'csv':
            tables = [pd.read_csv(filepath)]
        elif ext == 'pdf':
            tables = extract_tables_from_pdf(filepath)
        else:
            tables, ocr_error = extract_tables_from_image(filepath)
            if ocr_error:
                response = {'error': ocr_error}
                if is_image:
                    response['image_url'] = f'/api/image/{unique_name}'
                return jsonify(response), 400

        if not tables:
            response = {'error': 'No tables found in the uploaded file.'}
            if is_image:
                response['image_url'] = f'/api/image/{unique_name}'
            return jsonify(response), 400

        all_results = []
        grand_total = 0.0

        for i, df in enumerate(tables):
            sheet_name = f"Table {i + 1}"
            result = process_table(df, sheet_name)
            all_results.append(result)
            if result.get('summary', {}).get('content_sum') is not None:
                grand_total += result['summary']['content_sum']

        response = {
            'filename': file.filename,
            'total_tables': len(all_results),
            'grand_total_content': round(grand_total, 4),
            'tables': all_results
        }
        if is_image:
            response['image_url'] = f'/api/image/{unique_name}'

        return jsonify(response)

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        # Keep images for preview, clean up others
        if not is_image and os.path.exists(filepath):
            os.remove(filepath)


@app.route('/api/image/<filename>')
def serve_image(filename):
    """Serve uploaded images for preview."""
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/verify', methods=['POST'])
def verify_manual():
    """Verify manually entered table data."""
    data = request.get_json()
    if not data or 'rows' not in data:
        return jsonify({'error': 'No data provided'}), 400

    rows = data['rows']
    if not rows:
        return jsonify({'error': 'Empty table'}), 400

    # Build DataFrame from manual input
    df = pd.DataFrame(rows)
    result = process_table(df, "Manual Entry")

    response = {
        'filename': 'Manual Entry',
        'total_tables': 1,
        'grand_total_content': result.get('summary', {}).get('content_sum', 0),
        'tables': [result]
    }
    return jsonify(response)


@app.route('/')
def index():
    return send_from_directory(FRONTEND_DIR, 'index.html')


# ── Account management ────────────────────────────────────────────

def _load_accounts():
    if os.path.isfile(ACCOUNTS_FILE):
        with open(ACCOUNTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def _save_accounts(accounts):
    with open(ACCOUNTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(accounts, f, ensure_ascii=False, indent=2)


def _hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()


@app.route('/api/account/register', methods=['POST'])
def register_account():
    data = request.get_json()
    account = _safe_name(data.get('account', ''))
    password = data.get('password', '')
    if not account or not password:
        return jsonify({'error': 'Account name and password are required'}), 400
    if len(password) < 4:
        return jsonify({'error': 'Password must be at least 4 characters'}), 400

    accounts = _load_accounts()
    if account in accounts:
        return jsonify({'error': 'Account already exists. Please login.'}), 409

    accounts[account] = _hash_password(password)
    _save_accounts(accounts)
    os.makedirs(os.path.join(SAVED_BILLS_DIR, account), exist_ok=True)
    return jsonify({'ok': True})


@app.route('/api/account/login', methods=['POST'])
def login_account():
    data = request.get_json()
    account = _safe_name(data.get('account', ''))
    password = data.get('password', '')
    if not account or not password:
        return jsonify({'error': 'Account name and password are required'}), 400

    accounts = _load_accounts()
    if account not in accounts:
        return jsonify({'error': 'Account not found. Please register first.'}), 404
    if accounts[account] != _hash_password(password):
        return jsonify({'error': 'Incorrect password.'}), 401

    return jsonify({'ok': True})


# ── Bill save / load / list / delete ──────────────────────────────

def _safe_name(name):
    """Sanitize a user-supplied name to a safe filesystem component."""
    return re.sub(r'[^A-Za-z0-9_\- ]', '', str(name).strip())[:120]


@app.route('/api/bills/save', methods=['POST'])
def save_bill():
    """Save a bill's dashboard data as JSON under saved_bills/<account>/."""
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({'error': 'No data received'}), 400

    account = _safe_name(data.get('account', 'default')) or 'default'
    sheet = _safe_name(data.get('sheetName', 'Untitled')) or 'Untitled'
    bill = _safe_name(data.get('billNumber', 'NA')) or 'NA'
    date = _safe_name(data.get('billDate', ''))

    acct_dir = os.path.join(SAVED_BILLS_DIR, account)
    os.makedirs(acct_dir, exist_ok=True)

    filename = f"{sheet}_{bill}_{date}.json"
    filepath = os.path.join(acct_dir, filename)

    payload = {
        'sheetName': data.get('sheetName', ''),
        'billNumber': data.get('billNumber', ''),
        'billDate': data.get('billDate', ''),
        'rows': data.get('rows', [])
    }
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return jsonify({'ok': True, 'filename': filename})


@app.route('/api/bills/list/<account>', methods=['GET'])
def list_bills(account):
    """List all saved bills for a given account."""
    account = _safe_name(account)
    acct_dir = os.path.join(SAVED_BILLS_DIR, account)
    if not os.path.isdir(acct_dir):
        return jsonify({'bills': []})

    bills = []
    for fname in sorted(os.listdir(acct_dir)):
        if fname.endswith('.json'):
            fpath = os.path.join(acct_dir, fname)
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    info = json.load(f)
                bills.append({
                    'filename': fname,
                    'sheetName': info.get('sheetName', ''),
                    'billNumber': info.get('billNumber', ''),
                    'billDate': info.get('billDate', ''),
                })
            except Exception:
                bills.append({'filename': fname})
    return jsonify({'bills': bills})


@app.route('/api/bills/load/<account>/<filename>', methods=['GET'])
def load_bill(account, filename):
    """Load a specific saved bill."""
    account = _safe_name(account)
    filename = _safe_name(filename.replace('.json', '')) + '.json'
    fpath = os.path.join(SAVED_BILLS_DIR, account, filename)
    if not os.path.isfile(fpath):
        return jsonify({'error': 'Bill not found'}), 404
    with open(fpath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return jsonify(data)


@app.route('/api/bills/delete/<account>/<filename>', methods=['DELETE'])
def delete_bill(account, filename):
    """Delete a saved bill."""
    account = _safe_name(account)
    filename = _safe_name(filename.replace('.json', '')) + '.json'
    fpath = os.path.join(SAVED_BILLS_DIR, account, filename)
    if os.path.isfile(fpath):
        os.remove(fpath)
        return jsonify({'ok': True})
    return jsonify({'error': 'Not found'}), 404


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--prod', action='store_true', help='Run with waitress (production)')
    parser.add_argument('--port', type=int, default=5000)
    args = parser.parse_args()

    if args.prod:
        from waitress import serve
        print(f'Serving on http://0.0.0.0:{args.port}  (accessible to everyone on the network)')
        serve(app, host='0.0.0.0', port=args.port, threads=8)
    else:
        app.run(debug=True, port=args.port)
