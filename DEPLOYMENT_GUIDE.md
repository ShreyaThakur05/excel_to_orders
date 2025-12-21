# Order Automation - Deployment Guide

## 📦 Files to Copy to New System

Copy these files to the new computer:
```
order_automation.py
config.json
requirements.txt
run_order_automation.bat
service_account.json  (your Google Cloud key)
```

## 🔧 Setup on New System

### 1. Install Python 3.8+
Download from: https://python.org

### 2. Install Dependencies
```cmd
pip install -r requirements.txt
```

### 3. Verify Files
- ✅ `service_account.json` (Google Cloud key)
- ✅ `config.json` (your settings)
- ✅ `order_automation.py` (main script)

### 4. Test Run
```cmd
python order_automation.py
```

## 🚀 Ready to Use

Double-click `run_order_automation.bat` or run:
```cmd
python order_automation.py
```

## 📝 Notes
- Same Google Sheet works on any system
- Same Shoonya credentials work anywhere
- Email notifications work from any computer
- No additional setup needed

## 🔒 Security
- Keep `service_account.json` secure
- Don't share `config.json` (contains credentials)
- Use on trusted computers only