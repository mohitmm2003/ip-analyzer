
# IP Analyzer

This Python script, named `ip_analyzer.py`, is designed to help you identify and analyze potential abuse of IP addresses by querying the AbuseIPDB API. The script reads IP addresses from an input Excel file (`input_ips.xlsx`), checks each IP for abuse reports, and saves the results in a detailed output Excel file with a timestamp in its name (e.g., `output_results_2024-07-21_12-30-00.xlsx`).

Key features include converting report times to the Asia/Kolkata timezone, ensuring you have the relevant time context for each report. The script gathers comprehensive information, such as the public status of the IP, IP version, whitelist status, confidence of abuse score, country code, usage type, ISP, domain name, hostnames, Tor usage status, total reports, and the last reported time.

The results are centrally aligned and formatted for easy readability. The script incorporates a delay between API requests to prevent exceeding rate limits. This tool is essential for cybersecurity analysts and network administrators looking to monitor and manage the reputation of IP addresses in their networks effectively. It automates the process of gathering critical abuse data, saving time and effort while providing accurate and timely information.

## Features
- Reads IP addresses from an input Excel file.
- Queries AbuseIPDB for details about each IP address.
- Converts report times to Asia/Kolkata timezone.
- Saves the results to an output Excel file with formatted columns.

## Prerequisites
- Python 3.x
- Required Python packages:
  - pandas
  - requests
  - pytz
  - openpyxl

## Installation

1. Clone the repository or download the script.
2. Install the required Python packages:
   ```bash
   pip install pandas requests pytz openpyxl
   ```

## Usage

1. Prepare your input file:
   - Create an Excel file named `input_ips.xlsx`.
   - Ensure it has a column named `IP` containing the IP addresses you want to check.

2. Update the API key:
   - Replace the placeholder `API_KEY` in the script with your AbuseIPDB API key.

3. Run the script:
   ```bash
   python ip_abuse_checker.py
   ```

4. The results will be saved to an output Excel file with a timestamp in its name, e.g., `output_results_2024-07-21_12-30-00.xlsx`.

## Code Explanation

- **Reading the input file:**
  ```python
  df = pd.read_excel(INPUT_FILE)
  ips = df['IP'].tolist()
  ```

- **Querying AbuseIPDB:**
  ```python
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
  ```

- **Converting time to Asia/Kolkata timezone:**
  ```python
  def convert_to_kolkata_time(iso_time_str):
      if iso_time_str:
          utc_time = datetime.fromisoformat(iso_time_str.replace('Z', '+00:00'))
          kolkata_time = utc_time.astimezone(KOLKATA_TZ)
          return kolkata_time.strftime('%Y-%m-%d %H:%M:%S')
      return ''
  ```

- **Processing each IP and saving results:**
  ```python
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
      time.sleep(1.5)
  ```

- **Saving results to Excel:**
  ```python
  results_df.to_excel(OUTPUT_FILE, index=False)
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
              cell.alignment = Alignment(horizontal='center', vertical='center')
          adjusted_width = (max_length + 2)
          worksheet.column_dimensions[column_name].width = adjusted_width
  ```

## Notes
- Ensure your IP column is named `IP` in the input Excel file.
- The script includes a delay (`time.sleep(1.5)`) between API requests to prevent hitting rate limits.
