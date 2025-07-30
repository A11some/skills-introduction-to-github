# Outlook Email Search Tool

This Python script searches your Outlook inbox for emails containing the word "beat" in the title or content from the start of the current month until today. It generates both an HTML report and a text file with the results.

## Features

- 🔍 Search emails by keyword in subject or body
- 📅 Filter by date range (start of month to today)
- 🎨 Generate beautiful HTML report with highlighting
- 📄 Save results to text file
- 🔐 Secure connection to Outlook/Exchange
- 🏢 Support for both personal and corporate Exchange servers

## Prerequisites

- Python 3.6 or higher
- Outlook/Exchange account credentials
- Network access to your Exchange server

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **For corporate environments (if needed):**
   - You may need to configure your Exchange server URL
   - Contact your IT department for the correct server settings

## Usage

1. **Run the script:**
   ```bash
   python outlook_email_search.py
   ```

2. **Enter your credentials when prompted:**
   - Email address
   - Password
   - Exchange server URL (optional - leave empty for autodiscover)

3. **The script will:**
   - Connect to your Outlook account
   - Search for emails containing "beat" from the start of the month to today
   - Generate an HTML report (`outlook_search_results.html`)
   - Save a text version (`outlook_search_results.txt`)

## Output Files

### HTML Report (`outlook_search_results.html`)
- Beautiful, formatted display of all matching emails
- Search term highlighting
- Responsive design
- Summary information
- Easy to read and share

### Text File (`outlook_search_results.txt`)
- Plain text format
- Complete email content
- Structured layout
- Easy to search and process further

## Configuration Options

### Changing the Search Term
To search for a different word instead of "beat", modify line 245 in the script:
```python
search_term = "your_search_term_here"
```

### Changing the Date Range
To modify the date range, edit the `get_start_of_month()` function or modify lines 246-247:
```python
start_date = get_start_of_month()  # or your custom date
end_date = date.today()            # or your custom date
```

## Troubleshooting

### Connection Issues
- **"Authentication failed"**: Check your email and password
- **"Server not found"**: Enter the correct Exchange server URL
- **SSL errors**: The script includes SSL verification bypass for corporate environments

### No Emails Found
- Verify the search term exists in your emails
- Check the date range is correct
- Ensure you have access to the inbox

### Permission Issues
- Make sure your account has delegate access to the inbox
- Contact your IT department if you need elevated permissions

## Security Notes

- The script stores credentials only in memory during execution
- No credentials are saved to files
- SSL verification can be disabled for corporate environments
- Always use secure connections when possible

## Example Output

The HTML report will show:
- Search summary with date range and count
- Each email with subject, sender, date, and body
- Highlighted search terms in the email content
- Professional styling for easy reading

The text file will contain:
- Plain text version of all emails
- Structured format for easy processing
- Complete email metadata

## Support

If you encounter issues:
1. Check your network connection
2. Verify your Exchange server settings
3. Ensure you have the correct permissions
4. Try running with verbose logging for debugging

## License

This script is provided as-is for educational and personal use. Please ensure you comply with your organization's email usage policies.