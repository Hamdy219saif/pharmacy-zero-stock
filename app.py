from flask import Flask, request, send_file, render_template
import pandas as pd
import io

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

    # Read files
    warehouse = pd.read_excel(warehouse_file)
    branch = pd.read_excel(branch_file)

    # Clean column names
    warehouse.columns = warehouse.columns.str.strip()
    branch.columns = branch.columns.str.strip()

    # Normalize codes to string
    warehouse['كود الصنف'] = warehouse['كود الصنف'].astype(str).str.strip()
    branch['رقم الصنف'] = branch['رقم الصنف'].astype(str).str.strip()

    # Items available in warehouse (stock > 0)
    warehouse_available = warehouse[warehouse['رصيد المستودع'] > 0]

    # Codes present at branch
    branch_codes = set(branch['رقم الصنف'].unique())

    # Items with zero stock at branch = not in branch file
    zero_at_branch = warehouse_available[~warehouse_available['كود الصنف'].isin(branch_codes)]

    # Build result
    result = zero_at_branch[['كود الصنف', 'الصنف', 'رصيد المستودع']].copy()
    result.columns = ['كود الصنف', 'اسم الصنف', 'رصيد المستودع']
    result = result.sort_values('اسم الصنف').reset_index(drop=True)

    # Write to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name='الاحتياجات')

        # Style the sheet
        workbook = writer.book
        worksheet = writer.sheets['الاحتياجات']

        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        # Header style
        header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True, size=12)

        for col in range(1, 4):
            cell = worksheet.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        # Column widths
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 40
        worksheet.column_dimensions['C'].width = 18

        # Alternate row colors
        light_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
        for row in range(2, len(result) + 2):
            if row % 2 == 0:
                for col in range(1, 4):
                    worksheet.cell(row=row, column=col).fill = light_fill

    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='الاحتياجات_من_المستودع.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True, port=5050)
