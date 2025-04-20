HR Leave Request Automation System
==================================

This project automates the HR Leave Request process using Python. It simulates a structured workflow where employees submit leave requests, supervisors and HR review and approve/reject them, and all stakeholders receive email notifications.

Project Structure
-----------------

leave-request-automation/
├── leave_request_automation.py       # Main Python script
├── Leave_Request_Log.xlsx            # Excel log file (created/updated by script)
└── README.md                         # Instructions and documentation

Automation Workflow
-------------------

- Employee submits leave request on the platform
- The leave request is automatically logged on an excel spreadsheet
- Email notifications are sent to:
  - Supervisor upon submission
  - HR upon supervisor approval
  - Employee upon supervisor rejection
  - Employee after final HR decision
- The following status can be tracked on the platform: New → Pending HR → Approved/Rejected

Requirements
------------

- Python 3.7+
- Required Python packages:
  `pip install pandas openpyxl`

Email Setup
-----------

To enable email notifications, you'll need to:

1. Use a Gmail account (or other SMTP-compatible service)
2. Enable App Passwords in your Google account (if 2FA is on)
3. Update the following variables in the script:

```python
EMAIL_ADDRESS = 'your_email@gmail.com'
EMAIL_PASSWORD = 'your_app_password'
SUPERVISOR_EMAIL = 'supervisor_email@example.com'
HR_EMAIL = 'hr_email@example.com'

How to Run
----------

1. Open a terminal or command prompt.
2. Navigate to the project directory.
3. Run the script:
   python leave_request_automation.py

Follow the on-screen prompts to simulate the full HR leave workflow.

Notes
-----
- The script automatically creates or updates Leave_Request_Log.xlsx in the same directory.
- There might be some latency with email delivery.
- All decisions (Supervisor and HR) are manually entered for demo purposes.

