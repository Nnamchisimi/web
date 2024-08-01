from flask import Flask, request, render_template
import pandas as pd
import webbrowser
import threading
import time

app = Flask(__name__)

# Load the Excel file into a global DataFrame
file_path = r'C:\Users\User\Desktop\web\zcstcketrd_expstckentryd_status_trcklnk_trckno.xlsx'  # Correct file path
df = pd.read_excel(file_path)

# List of columns to display
columns_to_display = [
    'Part number',
    'Designation of part no. Ordered',
    'DN no.',
    'Inv. no.',
    'Q ordered',
    'Customer info',
    'Dist. ch.',
    'Items status',
    'tracking_details',
    'expectedstockentrydate',
    'Expected DD expecte to be dispatched',
    'status',
    'StockEntryDate'
]

@app.route('/', methods=['GET', 'POST'])
def index():
    part_details_html = None  # Initialize the variable
    no_details_found = False  # Flag to indicate if no details were found

    if request.method == 'POST':
        search_query = request.form.get('search_query')

        if search_query:
            search_query = search_query.strip()  # Ensure no leading or trailing spaces

            # Create filters based on the search query
            if 'backorder' in search_query.lower():
                filters = (
                    df['Customer info'].astype(str).str.contains(search_query, case=False, na=False) |
                    df['Part number'].astype(str).str.contains(search_query, case=False, na=False) |
                    (df['list_of_backorders'].astype(str).str.contains(search_query, case=False, na=False) &
                     (df['status'] != 'Rejected'))
                )
            else:
                filters = (
                    df['Customer info'].astype(str).str.contains(search_query, case=False, na=False) |
                    df['Part number'].astype(str).str.contains(search_query, case=False, na=False) |
                    df['list_of_backorders'].astype(str).str.contains(search_query, case=False, na=False) |
                    (df['status'].str.contains(search_query, case=False, na=False))
                )

            part_details = df[filters]
            if not part_details.empty:
                part_details = part_details[columns_to_display]  # Filter columns

                # Convert tracking_details to clickable links
                part_details['tracking_details'] = part_details['tracking_details'].apply(lambda x: f'<a href="{x}" target="_blank">{x}</a>' if pd.notnull(x) and x != 'Not Found' else '')

                # Convert DataFrame to HTML and allow for safe rendering
                part_details_html = part_details.to_html(classes='table table-striped', escape=False, index=False)
            else:
                no_details_found = True
        else:
            no_details_found = True

    return render_template('index.html', part_details=part_details_html, no_details_found=no_details_found)

def open_browser():
    """Open the browser to the Flask app URL."""
    url = 'http://127.0.0.1:5000/'
    webbrowser.open_new(url)

if __name__ == '__main__':
    # Start the Flask server in a new thread
    def run_server():
        app.run(debug=True, use_reloader=False)  # Turn off reloader to avoid multiple starts

    server_thread = threading.Thread(target=run_server)
    server_thread.start()

    # Allow some time for the server to start before opening the browser
    time.sleep(1)  # Adjust if necessary

    # Open the browser
    open_browser()
