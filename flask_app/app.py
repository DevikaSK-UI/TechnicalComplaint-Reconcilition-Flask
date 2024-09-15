from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

comparison_data = None
output_filename = None

def perform_comparison(ccc_file, edc_file):
    ccc_df = pd.read_excel(ccc_file)
    edc_df = pd.read_excel(edc_file)

    # Define column order
    ccc_columns = ['Subject/Patient ID', 'Technical Complaint No.', 'AE related', 'DUN Number', 'Trial/Study Number']
    edc_columns = ['Subject', 'Seq No', 'AE related', 'Dispense Unit Number ID', 'Trial/Study Number']

    ccc_df = ccc_df[ccc_columns]
    edc_df = edc_df[edc_columns]
    
    # Rename CCC columns to match EDC columns
    ccc_df = ccc_df.rename(columns={
        'Subject/Patient ID': 'Subject',
        'Technical Complaint No.': 'Seq No',
        'DUN Number': 'Dispense Unit Number ID'
    })

    # Initialize comparison columns
    edc_df['Status'] = 'Not Present'
    edc_df['Mismatch_Details'] = ''

    # Perform comparison
    for i, row in edc_df.iterrows():
        ccc_row = ccc_df.loc[ccc_df['Subject'] == row['Subject']]
        if not ccc_row.empty:
            mismatch_details = []
            if row['Seq No'] != ccc_row.iloc[0]['Seq No']:
                mismatch_details.append(f"Seq No (EDC: {row['Seq No']} / CCC: {ccc_row.iloc[0]['Seq No']})")
            if row['AE related'] != ccc_row.iloc[0]['AE related']:
                mismatch_details.append(f"AE related (EDC: {row['AE related']} / CCC: {ccc_row.iloc[0]['AE related']})")
            if row['Dispense Unit Number ID'] != ccc_row.iloc[0]['Dispense Unit Number ID']:
                mismatch_details.append(f"Dispense Unit Number ID (EDC: {row['Dispense Unit Number ID']} / CCC: {ccc_row.iloc[0]['Dispense Unit Number ID']})")
            if row['Trial/Study Number'] != ccc_row.iloc[0]['Trial/Study Number']:
                mismatch_details.append(f"Trial/Study Number (EDC: {row['Trial/Study Number']} / CCC: {ccc_row.iloc[0]['Trial/Study Number']})")
            
            if mismatch_details:
                edc_df.at[i, 'Mismatch_Details'] = "; ".join(mismatch_details)
                edc_df.at[i, 'Status'] = 'Mismatch'
            else:
                edc_df.at[i, 'Status'] = 'Match'
        else:
            edc_df.at[i, 'Status'] = 'Not Present'
            edc_df.at[i, 'Mismatch_Details'] = 'Not present in CCC'

    # Reorder the columns for the final output
    edc_df = edc_df[['Subject', 'Seq No', 'Dispense Unit Number ID', 'AE related', 'Trial/Study Number', 'Status', 'Mismatch_Details']]

    return edc_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload')
def upload():
    return render_template('upload.html')

@app.route('/compare', methods=['POST'])
def compare():
    global comparison_data, output_filename
    ccc_file = request.files['ccc_file']
    edc_file = request.files['edc_file']

    # Generate output filename
    output_filename = f"{datetime.now().strftime('%Y-%m-%d')}_Technical_Reconciliation.xlsx"

    # Perform comparison
    comparison_data = perform_comparison(ccc_file, edc_file)

    # Save the comparison result as an Excel file
    comparison_data.to_excel(output_filename, index=False)

    # Return success response with data for table display
    return jsonify({'status': 'success', 'file': output_filename, 'data': comparison_data.to_dict(orient='records')})

@app.route('/results')
def results():
    global comparison_data

    if comparison_data is None:
        return render_template('upload.html', message="No comparison data available.")

    # Convert the comparison data to an HTML table
    html_table = comparison_data.to_html(classes='table table-striped', index=False)

    return render_template('results.html', table=html_table)

@app.route('/download/<filename>')
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
