# -*- coding: utf-8 -*-
"""
Created on Fri Apr 18 15:09:57 2025

@author: user
"""

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os


###### Core configurations

EXCEL_PATH = 'C:/Users/user.DESKTOP-7HJ2O1A/OneDrive - National College of Ireland/Desktop/Leave Management/Leave_Request_Log.xlsx'
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = 'nikobpmn@gmail.com'  # Email address used to send notifications
EMAIL_PASSWORD = 'rbkk hary ryws atvw'

SUPERVISOR_EMAIL = 'ijanosyke@gmail.com'
HR_EMAIL = 'ijanosyke@yahoo.com'


###### Creating function: Send Email

def send_email(to_address, subject, body):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)


###### Loading or Creating the Leave Request Log Excel File

if os.path.exists(EXCEL_PATH):
    leave_log = pd.read_excel(EXCEL_PATH, dtype={'Request_ID': str, 'Supervisor_Approval': str, 'HR_Approval': str})
else:
    leave_log = pd.DataFrame(columns=[
        'Request_ID', 'Employee_Name', 'Email', 'Start_Date', 'End_Date', 'Reason',
        'Supervisor_Approval', 'HR_Approval', 'Status'
    ])


###### Step 1 - Employee Fill Leave Request Form Inputs

print("Please fill in your leave request details:\n")
employee_name = input("Employee Name: ")
email = input("Email: ")
start_date = input("Start Date (YYYY-MM-DD): ")
end_date = input("End Date (YYYY-MM-DD): ")

reasons = ['Annual Leave', 'Sick Leave', 'Leave of Absence', 'Maternity Leave', 'Family Emergency']
print("\nPlease select your reason for leave:")
for i, r in enumerate(reasons, 1):
    print(f"{i}. {r}")

while True:
    try:
        choice = int(input("Enter the number of your choice (1–5): "))
        if 1 <= choice <= len(reasons):
            reason = reasons[choice - 1]
            break
        else:
            print("Please enter a number between 1 and 5.")
    except ValueError:
        print("Invalid input. Please enter a number.")

new_id = str(int(leave_log['Request_ID'].max()) + 1 if not leave_log.empty else 1).zfill(3)

new_row = {
    'Request_ID': new_id,
    'Employee_Name': employee_name,
    'Email': email,
    'Start_Date': start_date,
    'End_Date': end_date,
    'Reason': reason,
    'Supervisor_Approval': '',
    'HR_Approval': '',
    'Status': 'New'
}

leave_log = pd.concat([leave_log, pd.DataFrame([new_row])], ignore_index=True)
leave_log.to_excel(EXCEL_PATH, index=False)
print(f"\nLeave request submitted successfully! Your Request ID is: {new_id}")


###### Step 2: Supervisor Notification and Action

row = leave_log[leave_log['Request_ID'] == new_id].iloc[0]
print(f"\nSupervisor Action Required for Request ID {row['Request_ID']} - {row['Employee_Name']}")
send_email(SUPERVISOR_EMAIL,
           f"Leave Request from {row['Employee_Name']}",
           f"""
Hello Supervisor,

{row['Employee_Name']} has requested leave from {row['Start_Date']} to {row['End_Date']} for the reason: "{row['Reason']}".

Please review the request for Request ID {row['Request_ID']}.

Thanks,
Leave Management System
""")

while True:
    supervisor_decision = input("Enter Supervisor decision (Approve/Reject): ").strip().lower()
    if supervisor_decision in ['approve', 'reject']:
        break
    else:
        print("Please enter either 'Approve' or 'Reject'.")

if supervisor_decision == 'approve':
    leave_log.loc[leave_log['Request_ID'] == new_id, 'Supervisor_Approval'] = 'Approved'
    leave_log.loc[leave_log['Request_ID'] == new_id, 'Status'] = 'Pending HR'
    print("Supervisor approved the request.")


###### Step 2: HR Notification and Action

    send_email(HR_EMAIL,
               f"HR Review Required for Leave Request {row['Request_ID']}",
               f"""
Hello HR,

Supervisor has approved the leave request of {row['Employee_Name']} from {row['Start_Date']} to {row['End_Date']}.

Please provide your final decision.

Thanks,
Leave Management System
""")

    print(f"\nHR Action Required for Request ID {row['Request_ID']} - {row['Employee_Name']}")
    while True:
        hr_decision = input("Enter HR decision (Approve/Reject): ").strip().lower()
        if hr_decision in ['approve', 'reject']:
            break
        else:
            print("Please enter either 'Approve' or 'Reject'.")

    if hr_decision == 'approve':
        leave_log.loc[leave_log['Request_ID'] == new_id, 'HR_Approval'] = 'Approved'
        leave_log.loc[leave_log['Request_ID'] == new_id, 'Status'] = 'Approved'
        send_email(row['Email'], "Your Leave Request Approved",
                   f"Dear {row['Employee_Name']},\n\nYour leave request from {row['Start_Date']} to {row['End_Date']} has been approved by both Supervisor and HR.\n\nEnjoy your time off!\n\nBest regards,\nLeave Management System")
        print("HR approved the request. Employee notified.")
    else:
        leave_log.loc[leave_log['Request_ID'] == new_id, 'HR_Approval'] = 'Rejected'
        leave_log.loc[leave_log['Request_ID'] == new_id, 'Status'] = 'Rejected'
        send_email(row['Email'], "Your Leave Request Rejected by HR",
                   f"Dear {row['Employee_Name']},\n\nUnfortunately, your leave request from {row['Start_Date']} to {row['End_Date']} has been rejected by HR.\n\n\n\nRegards,\nLeave Management System")
        print("HR rejected the request. Employee notified.")

else:
    leave_log.loc[leave_log['Request_ID'] == new_id, 'Supervisor_Approval'] = 'Rejected'
    leave_log.loc[leave_log['Request_ID'] == new_id, 'Status'] = 'Rejected'
    send_email(row['Email'], "Your Leave Request Rejected by Supervisor",
               f"Dear {row['Employee_Name']},\n\nYour leave request from {row['Start_Date']} to {row['End_Date']} was rejected by your supervisor.\n\n\n\nRegards,\nLeave Management System")
    print("Supervisor rejected the request. Employee notified.")


 ###### Updating the Leave Request Log Excel File

leave_log.to_excel(EXCEL_PATH, index=False)
print("Leave Request Process Complete! Leave Management Log Updated.")
