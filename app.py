from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)
CORS(app)

# Helper: Smart column match
def find_column_flexible(df, column_names):
    for name in column_names:
        for col in df.columns:
            if str(col).strip().lower() == name.strip().lower():
                return col
    return None

# Helper: Generate email from name
def generate_email_from_name(full_name, company_name="company"):
    if not isinstance(full_name, str):
        return ""
    name_parts = full_name.lower().strip().split()
    email_base = ".".join(name_parts)
    return f"{email_base}.{company_name.lower()}@gmail.com"

@app.route('/send-emails', methods=['POST'])
def process_timesheet():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xls', '.xlsx')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel file.'}), 400

    try:
        df = pd.read_excel(file)

        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]

        # Attempt to map columns
        column_mapping = {
            'employee_name': find_column_flexible(df, ['employee name', 'name']),
            'date': find_column_flexible(df, ['date']),
            'planned_task': find_column_flexible(df, ['planned task', 'task planned']),
            'planned_hours': find_column_flexible(df, ['planned hours', 'hours planned']),
            'actual_task': find_column_flexible(df, ['actual task', 'task actual']),
            'actual_hours': find_column_flexible(df, ['actual hours', 'hours actual']),
        }

        # Employee Email (optional)
        employee_email_col = find_column_flexible(df, [
            'employee email', 'employeeemail', 'email', 'emp email'
        ])
        column_mapping['employee_email'] = employee_email_col
        has_email_column = employee_email_col is not None

        # Only required columns must be present
        required_cols = ['employee_name', 'date', 'planned_task', 'planned_hours', 'actual_task', 'actual_hours']
        missing_cols = [col for col in required_cols if not column_mapping[col]]
        if missing_cols:
            return jsonify({
                'error': f"❌ Missing or unrecognized columns: {', '.join([col.replace('_', ' ').title() for col in missing_cols])}",
                'available_columns': list(df.columns)
            }), 400

        # Rename for simplicity
        df = df.rename(columns={v: k for k, v in column_mapping.items() if v})

        # Generate fallback emails if missing
        company_domain = "company"
        if not has_email_column or 'employee_email' not in df.columns:
            df['employee_email'] = df['employee_name'].apply(
                lambda name: generate_email_from_name(name, company_domain)
            )
            df['email_generated'] = True
        else:
            df['employee_email'] = df['employee_email'].fillna('')
            mask = df['employee_email'].astype(str).str.strip() == ''
            df.loc[mask, 'employee_email'] = df.loc[mask, 'employee_name'].apply(
                lambda name: generate_email_from_name(name, company_domain)
            )
            df['email_generated'] = mask

        # Group by employee
        grouped = df.groupby('employee_name')

        summaries = []

        for employee, group in grouped:
            email = group['employee_email'].iloc[0]
            planned_total = group['planned_hours'].sum()
            actual_total = group['actual_hours'].sum()
            details = []
            for _, row in group.iterrows():
                details.append({
                    'date': str(row['date'].date()) if not pd.isna(row['date']) else '',
                    'planned_task': row['planned_task'],
                    'planned_hours': row['planned_hours'],
                    'actual_task': row['actual_task'],
                    'actual_hours': row['actual_hours'],
                })
            summaries.append({
                'employee': employee,
                'email': email,
                'planned_total_hours': planned_total,
                'actual_total_hours': actual_total,
                'email_generated': bool(group['email_generated'].iloc[0]),
                'details': details
            })

        return jsonify({'summaries': summaries}), 200

    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)
