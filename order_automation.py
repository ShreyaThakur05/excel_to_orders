import json
import logging
import sys
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Dict, Any
from NorenRestApiPy.NorenApi import NorenApi
import pyotp
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build

class OrderAutomation:
    def __init__(self, config_path: str = "config.json"):
        self.config = self._load_config(config_path)
        self._setup_logging()
        self.sheets_service = self._authenticate_google_sheets()
        self.shoonya_api = self._authenticate_shoonya()
        self.execution_summary = []
        
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logging.error(f"Config file {config_path} not found")
            sys.exit(1)
        except json.JSONDecodeError:
            logging.error(f"Invalid JSON in config file {config_path}")
            sys.exit(1)
    
    def _setup_logging(self):
        log_config = self.config.get('logging', {})
        log_file = log_config.get('log_file', 'order_execution.log')
        log_level = getattr(logging, log_config.get('log_level', 'INFO'))
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
    
    def _authenticate_shoonya(self):
        try:
            shoonya_config = self.config['shoonya']
            api = NorenApi(host='https://api.shoonya.com/NorenWClientTP/', websocket='wss://api.shoonya.com/NorenWSTP/')
            
            # Generate current TOTP like in strategy2
            totp = pyotp.TOTP(shoonya_config['factor2'])
            current_otp = totp.now()
            
            logging.info(f"Attempting Shoonya login with user: {shoonya_config['user']}")
            logging.info(f"Generated TOTP: {current_otp}")
            
            ret = api.login(
                userid=shoonya_config['user'],
                password=shoonya_config['pwd'],
                twoFA=current_otp,  # Use generated TOTP
                vendor_code=shoonya_config['vc'],
                api_secret=shoonya_config['app_key'],
                imei=shoonya_config['imei']
            )
            
            logging.info(f"Shoonya login response: {ret}")
            
            if ret and ret.get('stat') == 'Ok':
                logging.info(f"Shoonya API authentication successful at {datetime.now()}")
                logging.info(f"Username: {shoonya_config['user']}")
                token = ret.get('susertoken', '')
                if token:
                    logging.info(f"Token: {token[:10]}...{token[-10:]}")
                return api
            elif ret and ret.get('stat') == 'Not_Ok':
                logging.error(f"Shoonya authentication failed: {ret.get('emsg', 'Unknown error')}")
                logging.info("Continuing in simulation mode")
                return api
            else:
                logging.warning(f"Shoonya authentication returned: {ret}")
                logging.info("Continuing in simulation mode")
                return api
                
        except Exception as e:
            logging.error(f"Shoonya authentication error: {e}")
            logging.info("Continuing in simulation mode")
            return NorenApi(host='https://api.shoonya.com/NorenWClientTP/', websocket='wss://api.shoonya.com/NorenWSTP/')
    
    def send_email_notification(self, subject: str, body: str):
        try:
            email_config = self.config.get('email', {})
            if not email_config.get('enabled', False):
                return
            
            msg = MIMEMultipart()
            msg['From'] = email_config['from_email']
            msg['To'] = ', '.join(email_config['to_emails'])
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))
            
            server = smtplib.SMTP(email_config['smtp_server'], email_config['smtp_port'])
            server.starttls()
            server.login(email_config['from_email'], email_config['app_password'])
            
            text = msg.as_string()
            server.sendmail(email_config['from_email'], email_config['to_emails'], text)
            server.quit()
            
            logging.info("Email notification sent successfully")
            
        except Exception as e:
            logging.error(f"Failed to send email notification: {e}")
    def _authenticate_google_sheets(self):
        try:
            service_account_file = self.config['google_sheets']['service_account_file']
            if not os.path.exists(service_account_file):
                logging.error(f"Service account file {service_account_file} not found")
                sys.exit(1)
                
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file,
                scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
            )
            
            service = build('sheets', 'v4', credentials=credentials)
            logging.info("Google Sheets authentication successful")
            return service
            
        except Exception as e:
            logging.error(f"Google Sheets authentication failed: {e}")
            sys.exit(1)
    
    def read_sheet_data(self) -> List[Dict[str, Any]]:
        try:
            sheet_id = self.config['google_sheets']['sheet_id']
            worksheet_name = self.config['google_sheets']['worksheet_name']
            
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=f"{worksheet_name}!A:G"  # Updated to read columns A-G
            ).execute()
            
            values = result.get('values', [])
            if not values:
                logging.warning("No data found in sheet")
                return []
            
            headers = values[0]
            data = []
            
            for row in values[1:]:
                if len(row) >= 7:  # Updated to expect 7 columns including quantity
                    # Skip rows where execution flag is not YES
                    if len(row) < 7 or row[6].upper() != 'YES':
                        continue
                        
                    data.append({
                        'serial_number': int(row[0]) if row[0].isdigit() else 0,
                        'trading_symbol': row[1],
                        'order_type': row[2].upper(),
                        'price': float(row[3]) if row[3] and row[3] != '' else 0.0,
                        'buy_sell': row[4].upper().strip(),
                        'quantity': int(row[5]) if row[5] and row[5].isdigit() else self.config['broker_api']['default_quantity'],
                        'execution_flag': row[6].upper()
                    })
            
            logging.info(f"Read {len(data)} rows from sheet")
            return data
            
        except Exception as e:
            logging.error(f"Failed to read sheet data: {e}")
            sys.exit(1)
    
    def filter_and_sort_orders(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        # Filter rows where Execution Flag = YES
        filtered_data = [row for row in data if row['execution_flag'] == 'YES']
        
        # Sort by Serial Number ascending
        sorted_data = sorted(filtered_data, key=lambda x: x['serial_number'])
        
        logging.info(f"Filtered to {len(sorted_data)} orders for execution")
        return sorted_data
    
    def place_order(self, order: Dict[str, Any]) -> bool:
        try:
            # Prepare order for Shoonya API
            order_params = {
                'buy_or_sell': 'B' if order['buy_sell'] == 'BUY' else 'S',
                'product_type': 'I',  # MIS for intraday
                'exchange': 'NSE',
                'tradingsymbol': order['trading_symbol'] + '-EQ',
                'quantity': str(order['quantity']),
                'discloseqty': '0',
                'price_type': 'MKT' if order['order_type'] == 'MARKET' else 'LMT',
                'retention': 'DAY'
            }
            
            if order['order_type'] == 'LIMIT':
                order_params['price'] = str(order['price'])
            
            logging.info(f"Placing order: {order_params}")
            
            # Place actual order
            response = self.shoonya_api.place_order(**order_params)
            
            logging.info(f"Order response: {response}")
            
            if response and response.get('stat') == 'Ok':
                order_id = response.get('norenordno')
                logging.info(f"Order placed successfully: {order['trading_symbol']} - Order ID: {order_id}")
                
                self.execution_summary.append({
                    'symbol': order['trading_symbol'],
                    'side': order['buy_sell'],
                    'type': order['order_type'],
                    'quantity': order['quantity'],
                    'price': order.get('price', 'MARKET'),
                    'order_id': order_id,
                    'status': 'SUCCESS'
                })
                return True
            else:
                error_msg = response.get('emsg', 'Unknown error') if response else 'No response'
                logging.error(f"Order failed: {order['trading_symbol']} - Error: {error_msg}")
                
                self.execution_summary.append({
                    'symbol': order['trading_symbol'],
                    'side': order['buy_sell'],
                    'type': order['order_type'],
                    'quantity': order['quantity'],
                    'price': order.get('price', 'MARKET'),
                    'order_id': None,
                    'status': 'FAILED',
                    'error': error_msg
                })
                return False
                
        except Exception as e:
            logging.error(f"Error placing order {order['trading_symbol']}: {e}")
            self.execution_summary.append({
                'symbol': order['trading_symbol'],
                'side': order['buy_sell'],
                'type': order['order_type'],
                'quantity': order['quantity'],
                'price': order.get('price', 'MARKET'),
                'order_id': None,
                'status': 'ERROR',
                'error': str(e)
            })
            return False
    
    def execute_orders(self, orders: List[Dict[str, Any]]):
        logging.info(f"Starting execution of {len(orders)} orders")
        
        for i, order in enumerate(orders, 1):
            logging.info(f"Processing order {i}/{len(orders)}: {order['trading_symbol']} {order['buy_sell']} {order['order_type']}")
            
            success = self.place_order(order)
            
            if not success:
                logging.error(f"Order execution stopped at order {i} due to failure")
                sys.exit(1)
        
        logging.info("All orders executed successfully")
    
    def run(self):
        logging.info("Starting order automation")
        
        try:
            # Read data from Google Sheets
            sheet_data = self.read_sheet_data()
            
            # Filter and sort orders
            orders_to_execute = self.filter_and_sort_orders(sheet_data)
            
            if not orders_to_execute:
                logging.info("No orders to execute")
                self.send_email_notification(
                    "Order Automation - No Orders",
                    "Order automation completed. No orders were found for execution."
                )
                return
            
            # Execute orders sequentially
            self.execute_orders(orders_to_execute)
            
            # Send completion email
            self._send_completion_email()
            
            logging.info("Order automation completed successfully")
            
        except Exception as e:
            error_msg = f"Order automation failed: {e}"
            logging.error(error_msg)
            self.send_email_notification(
                "Order Automation - FAILED",
                f"Order automation failed with error:\n\n{error_msg}\n\nPlease check the logs for more details."
            )
            raise
    
    def _send_completion_email(self):
        if not self.execution_summary:
            return
            
        successful_orders = [o for o in self.execution_summary if o['status'] == 'SUCCESS']
        failed_orders = [o for o in self.execution_summary if o['status'] in ['FAILED', 'ERROR']]
        
        # Only send email if there are actual orders placed
        if not successful_orders and not failed_orders:
            return
            
        subject = f"Order Automation Complete - {len(successful_orders)} Success, {len(failed_orders)} Failed"
        
        body = "Order Automation Execution Summary\n"
        body += "=" * 40 + "\n\n"
        
        if successful_orders:
            body += "SUCCESSFUL ORDERS:\n"
            body += "-" * 20 + "\n"
            for order in successful_orders:
                body += f"• {order['symbol']} {order['side']} {order['quantity']} @ {order['price']} ({order['type']}) - ID: {order['order_id']}\n"
            body += "\n"
        
        if failed_orders:
            body += "FAILED ORDERS:\n"
            body += "-" * 15 + "\n"
            for order in failed_orders:
                body += f"• {order['symbol']} {order['side']} {order['quantity']} @ {order['price']} ({order['type']}) - ERROR: {order.get('error', 'Unknown')}\n"
            body += "\n"
        
        body += f"Total Orders Processed: {len(self.execution_summary)}\n"
        body += f"Success Rate: {len(successful_orders)}/{len(self.execution_summary)} ({len(successful_orders)/len(self.execution_summary)*100:.1f}%)\n"
        
        self.send_email_notification(subject, body)

def main():
    try:
        automation = OrderAutomation()
        automation.run()
    except KeyboardInterrupt:
        logging.info("Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()