from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)

# Load the Excel file into a global DataFrame
file_path = r'C:\Users\User\Desktop\web\updated_StarOrderplus_Order_items (1).xlsx'  # Replace with your file path
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
    'Tracking No',
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
        else:
            part_details = pd.DataFrame(columns=columns_to_display)  # Empty DataFrame with specified columns if column not found

    return render_template('index.html', part_details=part_details)

if __name__ == '__main__':
    app.run(debug=True)
