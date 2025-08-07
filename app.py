from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Email configuration 
SMTP_SERVER = "smtp.gmail.com"  
SMTP_PORT = 587
SMTP_USERNAME = "your mailid"  
SMTP_PASSWORD = "app password"  

def find_column_flexible(df, column_patterns):
    """Find column with flexible matching"""
    df_columns_lower = [str(col).strip().lower().replace('_', '').replace(' ', '') for col in df.columns]
    
    for pattern in column_patterns:
        pattern_clean = pattern.strip().lower().replace('_', '').replace(' ', '')
        for i, col_clean in enumerate(df_columns_lower):
            if pattern_clean in col_clean or col_clean in pattern_clean:
                return df.columns[i]
    return None

def generate_email_from_name(full_name, company_name="company"):
    """Generate email from employee name"""
    if not isinstance(full_name, str) or not full_name.strip():
        return ""
    
    name_parts = full_name.lower().strip().split()
    email_base = ".".join(name_parts)
    return f"{email_base}.{company_name.lower()}@gmail.com"

def send_email(to_email, subject, body, from_email=SMTP_USERNAME):
    """Send email via SMTP"""
    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        
        return True, "Email sent successfully"
    except Exception as e:
        return False, str(e)

def generate_email_content(employee_data):
    """Generate email content for an employee"""
    planned_hours = employee_data['planned_total_hours']
    actual_hours = employee_data['actual_total_hours']
    completion_rate = (actual_hours / planned_hours * 100) if planned_hours > 0 else 0
    hours_difference = actual_hours - planned_hours
    
    completion_status = 'Excellent' if completion_rate >= 90 else 'Good' if completion_rate >= 70 else 'Needs Improvement'
    
    # Find task mismatches
    task_mismatches = []
    for detail in employee_data['details']:
        if (detail['planned_task'].lower().strip() != detail['actual_task'].lower().strip() and 
            detail['planned_task'].strip() != '' and detail['actual_task'].strip() != ''):
            task_mismatches.append(detail)
    
    # Get date range
    dates = [detail['date'] for detail in employee_data['details'] if detail['date']]
    date_range = f"{min(dates)} to {max(dates)}" if dates else "N/A"
    
    email_content = f"""Dear {employee_data['employee']},

Here's your weekly timesheet summary:

📊 SUMMARY:
- Planned Hours: {planned_hours}
- Actual Hours: {actual_hours}  
- Completion Rate: {completion_rate:.1f}% ({completion_status})
- Hours Difference: {'+' if hours_difference > 0 else ''}{hours_difference}

📅 Week Period: {date_range}

"""
    
    if task_mismatches:
        email_content += f"""⚠️ TASK MISMATCHES ({len(task_mismatches)}):
"""
        for mismatch in task_mismatches:
            email_content += f"""- {mismatch['date']}: Planned "{mismatch['planned_task']}" → Actual "{mismatch['actual_task']}"
"""
        email_content += """
Please review these mismatches and ensure accurate task logging in future timesheets.
"""
    else:
        email_content += """✅ All tasks were completed as planned. Great work!

"""
    
    email_content += """If you have any questions about this summary, please contact HR.

Best regards,
HR Team"""
    
    return email_content

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'message': 'Employee Timesheet Backend API',
        'status': 'running',
        'endpoints': {
            '/send-emails': 'POST - Process timesheet and send emails'
        }
    })

@app.route('/send-emails', methods=['POST'])
def process_timesheet():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400
        
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel file.'}), 400

    try:
        # Read Excel file
        df = pd.read_excel(file)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        print(f"Available columns: {list(df.columns)}")
        
        # Skip empty rows or header rows
        initial_rows = len(df)
        df = df.dropna(how='all')
        
        # Try to detect actual data start
        if len(df) > 0:
            # Check if first row might be a title row
            first_row_values = [str(val) for val in df.iloc[0].values if pd.notna(val)]
            if any('timesheet' in val.lower() or 'sample' in val.lower() or 'employee' in val.lower() for val in first_row_values):
                # Skip potential title rows
                for i in range(min(3, len(df))):
                    # Try using row i as header
                    test_df = pd.read_excel(file, header=i)
                    test_df.columns = [str(col).strip() for col in test_df.columns]
                    
                    # Check if this gives us valid column names
                    valid_cols = sum([
                        bool(find_column_flexible(test_df, ['employee name', 'name'])),
                        bool(find_column_flexible(test_df, ['date'])),
                        bool(find_column_flexible(test_df, ['planned task', 'task planned'])),
                        bool(find_column_flexible(test_df, ['planned hours', 'hours planned'])),
                        bool(find_column_flexible(test_df, ['actual task', 'task actual'])),
                        bool(find_column_flexible(test_df, ['actual hours', 'hours actual']))
                    ])
                    
                    if valid_cols >= 5:  # At least 5 out of 6 required columns found
                        df = test_df
                        print(f"Using header row {i}, found {valid_cols} valid columns")
                        break
        
        if len(df) == 0:
            return jsonify({'error': 'No data found in Excel file after cleaning'}), 400

        # Flexible column mapping
        column_mapping = {
            'employee_name': find_column_flexible(df, ['employee name', 'employeename', 'name']),
            'date': find_column_flexible(df, ['date']),
            'planned_task': find_column_flexible(df, ['planned task', 'plannedtask', 'task planned']),
            'planned_hours': find_column_flexible(df, ['planned hours', 'plannedhours', 'hours planned']),
            'actual_task': find_column_flexible(df, ['actual task', 'actualtask', 'task actual']),
            'actual_hours': find_column_flexible(df, ['actual hours', 'actualhours', 'hours actual']),
            'employee_email': find_column_flexible(df, ['employee email', 'employeeemail', 'email', 'emp email'])
        }
        
        print(f"Column mapping: {column_mapping}")

        # Check required columns
        required_cols = ['employee_name', 'date', 'planned_task', 'planned_hours', 'actual_task', 'actual_hours']
        missing_cols = [col for col in required_cols if not column_mapping[col]]
        
        if missing_cols:
            friendly_names = {
                'employee_name': 'Employee Name',
                'date': 'Date', 
                'planned_task': 'Planned Task',
                'planned_hours': 'Planned Hours',
                'actual_task': 'Actual Task',
                'actual_hours': 'Actual Hours'
            }
            missing_friendly = [friendly_names[col] for col in missing_cols]
            return jsonify({
                'error': f"Missing or unrecognized columns: {', '.join(missing_friendly)}",
                'available_columns': list(df.columns)
            }), 400

        # Rename columns for easier processing
        rename_map = {v: k for k, v in column_mapping.items() if v is not None}
        df = df.rename(columns=rename_map)
        
        # Handle missing email column
        has_email_column = column_mapping['employee_email'] is not None
        if not has_email_column:
            print("No email column found, generating emails from names")
            df['employee_email'] = df['employee_name'].apply(
                lambda name: generate_email_from_name(name, "company")
            )
            df['email_generated'] = True
        else:
            # Fill missing emails
            df['employee_email'] = df['employee_email'].fillna('')
            mask = df['employee_email'].astype(str).str.strip() == ''
            df.loc[mask, 'employee_email'] = df.loc[mask, 'employee_name'].apply(
                lambda name: generate_email_from_name(name, "company")
            )
            df['email_generated'] = mask

        # Convert hours to numeric
        df['planned_hours'] = pd.to_numeric(df['planned_hours'], errors='coerce').fillna(0)
        df['actual_hours'] = pd.to_numeric(df['actual_hours'], errors='coerce').fillna(0)
        
        # Remove rows with missing essential data
        df = df.dropna(subset=['employee_name'])
        df = df[df['employee_name'].astype(str).str.strip() != '']
        
        if len(df) == 0:
            return jsonify({'error': 'No valid employee records found after cleaning'}), 400

        # Group by employee
        summaries = []
        grouped = df.groupby('employee_name')

        for employee, group in grouped:
            email = group['employee_email'].iloc[0]
            planned_total = group['planned_hours'].sum()
            actual_total = group['actual_hours'].sum()
            email_was_generated = group['email_generated'].iloc[0] if 'email_generated' in group.columns else not has_email_column
            
            details = []
            for _, row in group.iterrows():
                date_str = ''
                if pd.notna(row['date']):
                    if hasattr(row['date'], 'date'):
                        date_str = str(row['date'].date())
                    else:
                        date_str = str(row['date'])
                
                details.append({
                    'date': date_str,
                    'planned_task': str(row['planned_task']) if pd.notna(row['planned_task']) else '',
                    'planned_hours': float(row['planned_hours']),
                    'actual_task': str(row['actual_task']) if pd.notna(row['actual_task']) else '',
                    'actual_hours': float(row['actual_hours']),
                })
            
            summaries.append({
                'employee': str(employee),
                'email': str(email),
                'planned_total_hours': float(planned_total),
                'actual_total_hours': float(actual_total),
                'email_generated': bool(email_was_generated),
                'details': details
            })

        # Send emails (optional - uncomment if you want to actually send emails)
        
        email_results = []
        for summary in summaries:
            subject = f"Weekly Timesheet Summary - {datetime.now().strftime('%Y-%m-%d')}"
            email_content = generate_email_content(summary)
            
            success, message = send_email(summary['email'], subject, email_content)
            email_results.append({
                'employee': summary['employee'],
                'email': summary['email'], 
                'sent': success,
                'message': message
            })
        
        return jsonify({
            'message': f'Processed {len(summaries)} employee summaries',
            'summaries': summaries,
            'email_results': email_results
        }), 200
        
        
        # For demo purposes, just return summaries without actually sending emails
        return jsonify({
            'message': f'Processed {len(summaries)} employee summaries. Emails would be sent in production mode.',
            'summaries': summaries
        }), 200

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/test', methods=['GET'])
def test_backend():
    return jsonify({
        'message': 'Backend is working!',
        'timestamp': datetime.now().isoformat()
    })

if __name__ == '__main__':
    print("Starting Employee Timesheet Backend...")
    print("Available endpoints:")
    print("- GET /: API information")
    print("- POST /send-emails: Process timesheet file")
    print("- GET /test: Test backend connectivity")
    app.run(host='0.0.0.0', port=5000, debug=True)

