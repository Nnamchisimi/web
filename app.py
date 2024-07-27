from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)

# Load the Excel file into a global DataFrame
file_path = r'C:\Users\User\Desktop\web\fupdated_StarOrderplus_Order_items_with_expected_stock_entry_date.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# List of columns to display
columns_to_display = [
    'Designation of part no. ordered',
    'Designation of part no. confirmed',
    'Ord. date',
    'Items status',
    'Inv. date dispatch date',
    'Conf. DD confirmed dispatch date',
    'Expected DD expecte to be dispatched',
    'DN no.',
    'Inv. no.',
    'Q ordered',
    'Customer info',
    'Part number',
    'Dist. ch.',
    'Tracking No',
    'tracking_details',
    'expectedstockentrydate',
    'status',
    'StockEntryDate'
]

@app.route('/', methods=['GET', 'POST'])
def index():
    part_details = None
    if request.method == 'POST':
        pn_ordered = request.form.get('Customer_info').strip()
        if 'Customer info' in df.columns:
            part_details = df[df['Customer info'].astype(str).str.contains(pn_ordered, case=False, na=False)]
            part_details = part_details[columns_to_display]  # Filter columns

            # Convert tracking_details to clickable links
            part_details['tracking_details'] = part_details['tracking_details'].apply(lambda x: f'<a href="{x}" target="_blank">{x}</a>' if pd.notnull(x) and x != 'Not Found' else '')

            # Convert DataFrame to HTML and allow for safe rendering
            part_details_html = part_details.to_html(classes='table table-striped', escape=False, index=False)
        else:
            part_details_html = '<p>Column "Customer info" does not exist in the Excel file.</p>'

    return render_template('index.html', part_details=part_details_html)

if __name__ == '__main__':
    app.run(debug=True)
