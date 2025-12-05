# Junk Mail Email Scanner

This program scans your junk mail folder on Amazon WorkMail (Exchange-compatible) and extracts all unique email addresses.

## Setup

1. Install the required package:
```bash
pip install exchangelib
```

Or use the requirements file:
```bash
pip install -r requirements.txt
```

## Configuration

You'll need:
- Your email address
- Your password (or app-specific password if 2FA is enabled)
- Your Amazon WorkMail server address

### Finding Your WorkMail Server Address

For Amazon WorkMail, the server format is typically:
```
ews.mail.<region>.awsapps.com
```

Common examples:
- `ews.mail.us-east-1.awsapps.com`
- `ews.mail.us-west-2.awsapps.com`
- `ews.mail.eu-west-1.awsapps.com`

You can find your specific server in your WorkMail settings or contact your administrator.

## Usage

Run the program:
```bash
python email.py
```

You'll be prompted for:
1. Your email address
2. Your password
3. Your Exchange/WorkMail server address

The program will:
- Connect to your email server
- Scan all messages in your Junk folder
- Extract email addresses from:
  - Sender addresses
  - Email subjects
  - Email bodies
- Display unique email addresses sorted by frequency
- Optionally save results to a text file

## Output

The program displays:
- Total number of unique email addresses found
- Each email address with the count of how many times it appeared
- Results are sorted by frequency (most common first)

## Troubleshooting

If you encounter connection issues:

1. **Wrong server address**: Verify your WorkMail server address
2. **Authentication failed**: If 2FA is enabled, you may need to create an app-specific password
3. **EWS not enabled**: Contact your administrator to ensure Exchange Web Services (EWS) is enabled for your account
4. **Firewall/VPN**: Make sure you can reach the server (no firewall blocking)

## Security Notes

- Never commit your password to version control
- Consider using environment variables for credentials:
  ```bash
  export EMAIL_USER="your.email@example.com"
  export EMAIL_PASS="your-password"
  ```
- Use app-specific passwords when possible
- The program processes emails locally and doesn't send data anywhere
