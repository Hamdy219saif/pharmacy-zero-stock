from flask import Flask, request, send_file, render_template
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    warehouse_file = request.files.get('warehouse')
    branch_file = request.files.get('branch')

    if not warehouse_file or not branch_file:
        return 'يرجى رفع الملفين', 400

    try:
        warehouse = pd.read_excel(warehouse_file)
        branch = pd.read_excel(branch_file)
    except:
        return 'خطأ في قراءة الملفات', 400

    # تنظيف الأعمدة
    warehouse.columns = warehouse.columns.astype(str).str.strip()
    branch.columns = branch.columns.astype(str).str.strip()

    # 🔥 تحديد عمود الكود (أرقام طويلة)
    def detect_code(df):
        best_col = None
        max_numeric = 0

        for col in df.columns:
            values = df[col].astype(str)
            numeric_count = values.str.replace('.', '').str.isnumeric().sum()

            if numeric_count > max_numeric:
                max_numeric = numeric_count
                best_col = col

        return best_col

    # 🔥 تحديد عمود الكمية
    def detect_qty(df):
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                return col
        return None

    # 🔥 تحديد اسم الصنف (فيه حروف)
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

    # تحديد الأعمدة
    w_code = detect_code(warehouse)
    b_code = detect_code(branch)
    qty_col = detect_qty(warehouse)
    name_col = detect_name(warehouse, w_code)

    if not w_code or not b_code:
        return "❌ لم يتم تحديد عمود الكود"

    if not qty_col:
        return "❌ لم يتم تحديد عمود الكمية"

    if not name_col:
        name_col = warehouse.columns[0]

    # تجهيز البيانات
    warehouse[w_code] = warehouse[w_code].astype(str)
    branch[b_code] = branch[b_code].astype(str)

    # فلترة
    warehouse_available = warehouse[warehouse[qty_col] > 0]
    branch_codes = set(branch[b_code].unique())

    zero_at_branch = warehouse_available[
        ~warehouse_available[w_code].isin(branch_codes)
    ]

    # النتيجة
    result = zero_at_branch[[w_code, name_col, qty_col]].copy()
    result.columns = ['كود الصنف', 'اسم الصنف', 'الكمية']
    result = result.sort_values('اسم الصنف').reset_index(drop=True)

    # إخراج Excel
    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)

    filename = f'الاحتياجات_{datetime.now().strftime("%Y-%m-%d")}.xlsx'

    return send_file(output, download_name=filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
