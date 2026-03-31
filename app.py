from flask import Flask, request, send_file, render_template
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

def smart_read_excel(file):
    df = pd.read_excel(file, header=None)
    header_row = 0
    for i, row in df.iterrows():
        text = ' '.join(row.astype(str))
        if 'كود' in text or 'رقم' in text:
            header_row = i
            break
    df = pd.read_excel(file, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    return df

def detect_code(df):
    best_col = None
    max_numeric = 0
    for col in df.columns:
        values = df[col].astype(str)
        count = values.str.replace('.', '').str.isnumeric().sum()
        if count > max_numeric:
            max_numeric = count
            best_col = col
    return best_col

def detect_qty(df, code_col=None):
    for col in df.columns:
        if any(k in str(col) for k in ['رصيد', 'كمية', 'qty', 'stock', 'balance']):
            return col
    for col in df.columns:
        if col == code_col:
            continue
        nums = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(nums) > 5 and nums.median() < 100000:
            return col
    return None

def detect_name(df, code_col):
    best_col = None
    max_text = 0
    for col in df.columns:
        if col == code_col:
            continue
        values = df[col].astype(str)
        text_count = values.str.contains('[A-Za-zأ-ي]').sum()
        if text_count > max_text:
            max_text = text_count
            best_col = col
    return best_col

@app.route('/process', methods=['POST'])
def process():
    try:
        warehouse_file = request.files.get('warehouse')
        branch_file = request.files.get('branch')

        if not warehouse_file or not branch_file:
            return '❌ لازم ترفع الملفين'

        warehouse = smart_read_excel(warehouse_file)
        branch = smart_read_excel(branch_file)

        w_code = detect_code(warehouse)
        b_code = detect_code(branch)
        qty_col = detect_qty(warehouse, w_code)
        name_col = detect_name(warehouse, w_code)

        if not w_code or not b_code:
            return "❌ مش لاقي كود الصنف"
        if not qty_col:
            return "❌ مش لاقي عمود الكمية"
        if not name_col:
            name_col = warehouse.columns[0]

        warehouse[w_code] = warehouse[w_code].astype(str)
        branch[b_code] = branch[b_code].astype(str)

        warehouse_available = warehouse[pd.to_numeric(warehouse[qty_col], errors='coerce') > 0]
        branch_codes = set(branch[b_code].unique())

        result = warehouse_available[
            ~warehouse_available[w_code].isin(branch_codes)
        ][[w_code, name_col, qty_col]].copy()

        result.columns = ['كود الصنف', 'اسم الصنف', 'الكمية']

        output = io.BytesIO()
        result.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name=f"الاحتياجات_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"🔥 حصل خطأ: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
