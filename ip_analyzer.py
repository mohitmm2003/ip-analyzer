import pandas as pd
import requests
import time
from datetime import datetime
import pytz
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Data Constants
API_KEY = 'e3cefae33c33a8cb593960a6d6c812f028f96afbcf4f5c61f16dfbe539ae32c0bddaa0f6a50c7698'
INPUT_FILE = 'input_ips.xlsx'
ABUSEIPDB_URL = 'https://api.abuseipdb.com/api/v2/check'
KOLKATA_TZ = pytz.timezone('Asia/Kolkata')

# File Format
current_time = datetime.now(KOLKATA_TZ).strftime('%Y-%m-%d_%H-%M-%S')
OUTPUT_FILE = f'output_results_{current_time}.xlsx'

# Read the Excel file
df = pd.read_excel(INPUT_FILE)

# Ensure the IP column is named 'IP' in your Excel file
ips = df['IP'].tolist()

# Initialize list to store results
results = []

# Function to query AbuseIPDB
def query_abuseipdb(ip):
    headers = {
        'Key': API_KEY,
        'Accept': 'application/json'
    }
    params = {
        'ipAddress': ip,
        'maxAgeInDays': '90'
    }
    response = requests.get(ABUSEIPDB_URL, headers=headers, params=params)
    return response.json()

# Function to convert time to Asia/Kolkata timezone
def convert_to_kolkata_time(iso_time_str):
    if iso_time_str:
        utc_time = datetime.fromisoformat(iso_time_str.replace('Z', '+00:00'))
        kolkata_time = utc_time.astimezone(KOLKATA_TZ)
        return kolkata_time.strftime('%Y-%m-%d %H:%M:%S')
    return ''

# Loop through IPs and query AbuseIPDB
for ip in ips:
    result = query_abuseipdb(ip)
    data = result.get('data', {})
    last_reported_at = data.get('lastReportedAt', '')
    kolkata_time = convert_to_kolkata_time(last_reported_at)
    results.append({
        'IP Address': ip,
        'Is Public': data.get('isPublic', ''),
        'IP Version': data.get('ipVersion', ''),
        'Is Whitelisted': data.get('isWhitelisted', ''),
        'Confidence of Abuse': data.get('abuseConfidenceScore', ''),
        'Country Code': data.get('countryCode', ''),
        'Usage Type': data.get('usageType', ''),
        'ISP': data.get('isp', ''),
        'Domain Name': data.get('domain', ''),
        'Hostname(s)': ", ".join(data.get('hostnames', [])),
        'Is Tor': data.get('isTor', ''),
        'Total Reports': data.get('totalReports', ''),
        'Last Reported At': last_reported_at,
        'Last Reported At (Kolkata)': kolkata_time
    })
    # To prevent hitting rate limits
    time.sleep(1.5)

# Convert results to DataFrame
results_df = pd.DataFrame(results)

# Save the results to an Excel file
results_df.to_excel(OUTPUT_FILE, index=False)

# Adjust column widths and apply central alignment
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a') as writer:
    workbook = writer.book
    worksheet = workbook.active
    
    for column in worksheet.columns:
        max_length = 0
        column_name = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
            # Set cell alignment to center
            cell.alignment = Alignment(horizontal='center', vertical='center')
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_name].width = adjusted_width

print("Scanning complete. Results saved to", OUTPUT_FILE)
