# Order Automation Setup Guide

## Prerequisites
1. Python 3.8+ installed
2. Google Cloud Project with Sheets API enabled
3. Broker API credentials

## Google Sheets Setup

### 1. Create Google Cloud Service Account
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable Google Sheets API and Google Drive API
4. Go to IAM & Admin > Service Accounts
5. Create new service account
6. Download JSON key file as `service_account.json`

### 2. Share Google Sheet
1. Open your Google Sheet
2. Click Share button
3. Add the service account email (from JSON file) with Editor access
4. Copy the Sheet ID from URL (between `/d/` and `/edit`)

### 3. Sheet Format
Your Google Sheet must have these columns (A-G):
- A: Serial Number (1, 2, 3...)
- B: Trading Symbol (AAPL, MSFT, etc.)
- C: Order Type (MARKET or LIMIT)
- D: Price (required for LIMIT orders)
- E: Buy/Sell (BUY or SELL)
- F: Quantity (number of shares/lots)
- G: Execution Flag (YES or NO)

## Configuration

### 1. Update config.json
```json
{
    "google_sheets": {
        "sheet_id": "YOUR_ACTUAL_SHEET_ID",
        "worksheet_name": "Sheet1",
        "service_account_file": "service_account.json"
    },
    "broker_api": {
        "type": "shoonya",
        "default_quantity": 75
    },
    "shoonya": {
        "user": "YOUR_USER_ID",
        "pwd": "YOUR_PASSWORD",
        "factor2": "YOUR_2FA_TOKEN",
        "vc": "YOUR_VENDOR_CODE",
        "app_key": "YOUR_APP_KEY",
        "imei": "YOUR_IMEI"
    },
    "email": {
        "enabled": true,
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "from_email": "your-email@gmail.com",
        "app_password": "your-app-password",
        "to_emails": [
            "recipient@gmail.com"
        ]
    }
}
```

### 2. Place Files
- `service_account.json` in the same folder as the executable
- `config.json` with your actual settings

## Building the Executable

### Option 1: Use build script
```cmd
build.bat
```

### Option 2: Manual build
```cmd
pip install -r requirements.txt
pyinstaller order_automation.spec
```

## Running the Executable
1. Double-click `OrderAutomation.exe`
2. Check `order_execution.log` for results
3. Orders execute sequentially based on Serial Number
4. Only rows with Execution Flag = YES are processed
5. Email notifications sent on completion/failure

## Email Notifications
- **Success**: Summary of all executed orders with order IDs
- **Failure**: Error details and failed orders
- **No Orders**: Notification when no orders found for execution

## Troubleshooting
- Ensure service account JSON file is present
- Verify Google Sheet is shared with service account email
- Check Shoonya API credentials and 2FA token
- Verify email app password for Gmail
- Review log file for detailed error messages

## Security Notes
- Keep service account JSON file secure
- Don't commit API keys to version control
- Use environment variables for sensitive data in production