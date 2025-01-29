# PowerShell SFTP File Transfer Script

A robust PowerShell script for automated file transfers to SFTP servers using FileZilla Pro CLI, with email notifications and comprehensive logging.

## Features

- Automated file uploads to SFTP servers using FileZilla Pro CLI
- Comprehensive logging with multiple output options
- Email notifications with detailed transfer results
- File organization with success/failure sorting
- Environment variable configuration
- Secure credential handling
- Error handling and retry logic

## Prerequisites

### FileZilla Pro CLI

1. Purchase and download FileZilla Pro CLI from the official website at https://filezillapro.com/cli/
2. Install FileZilla Pro CLI to your preferred location (default: `C:\Program Files\FileZilla Pro CLI\`)
3. Ensure the installation path matches the `FILEZILLA_PATH` in your `.env` file

### PowerShell Requirements

- PowerShell 5.1 or higher
- System.Web assembly (automatically loaded by the script)

## Setup

1. Clone or download this repository
2. Copy `.env.example` to `.env`:
   ```powershell
   Copy-Item .env.example .env
   ```
3. Edit `.env` with your specific configuration:
   - Update all file paths
   - Configure SFTP settings
   - Set up email notifications (if desired)
   - Adjust debug and logging options

### Environment Variables

#### Debug Options
- `DEBUG`: Enable/disable debug logging
- `LOG_TO_FILE`: Enable/disable file logging
- `SEND_EMAIL`: Enable/disable email notifications
- `UPLOAD_TO_SFTP`: Enable/disable SFTP uploads

#### File Paths
- `FILE_PATH`: Input directory for files to transfer
- `FILE_PATH_PROCESSED`: Directory for successfully processed files
- `FILE_PATH_FAILED`: Directory for failed transfers
- `LOG_PATH`: Directory for log files

#### SFTP Configuration
- `FILEZILLA_PATH`: Path to FileZilla Pro CLI executable
- `SFTP_SERVER_HOST`: SFTP server hostname or IP
- `SFTP_SERVER_PORT`: SFTP server port (default: 22)
- `SFTP_USERNAME`: SFTP username
- `SFTP_PASSWORD`: SFTP password
- `SFTP_UPLOAD_PATH`: Remote directory path

#### Email Settings
- `EMAIL_TO`: Recipient email address
- `EMAIL_FROM`: Sender email address
- `EMAIL_SUBJECT`: Email notification subject
- `SMTP_HOST`: SMTP server hostname
- `SMTP_USERNAME`: SMTP username
- `SMTP_PASSWORD`: SMTP password
- `SMTP_TTLS`: Enable/disable TTLS (True/False)
- `SMTP_PORT`: SMTP port number

## Security Considerations

### SFTP Certificate Verification

By default, the script auto-accepts SFTP server certificates. For enhanced security:

1. Connect to the server using FileZilla Pro UI first
2. Export the server's certificate
3. Modify the connect command in the script to use the certificate:
   ```powershell
   connect --user $($env:SFTP_USERNAME) --pass $($env:SFTP_PASSWORD) --certfile "path/to/certificate" sftp://$($env:SFTP_SERVER_HOST):$($env:SFTP_SERVER_PORT)
   ```

### Password Security
- Store the `.env` file securely
- Consider using encrypted credentials or certificate-based authentication
- Restrict access to the script and configuration files

## Core Functions

### Load-EnvVariables
- Loads environment variables from `.env` file
- Handles path normalization
- Processes special cases for paths and sensitive data

### Write-Log
- Manages logging to console, file, and email
- Supports multiple log levels (INFO, DEBUG, ERROR)
- Timestamps all entries

### Send-Email
- Sends HTML-formatted email notifications
- Includes script output and execution results
- Supports TLS/SSL
- Handles authentication

### Upload-FileWithFileZilla
- Manages file transfers using FileZilla Pro CLI
- Creates and manages temporary script files
- Handles output and error streams
- Implements timeout and cleanup

### Process-Files
- Orchestrates the file transfer workflow
- Creates required directories
- Moves files to appropriate directories based on transfer success/failure

## Running the Script

1. Ensure all prerequisites are installed
2. Configure your `.env` file
3. Run the script:
   ```powershell
   .\Transfer_to_SFTP.ps1
   ```

## Logging

The script provides three types of logging:
1. Console output
2. File logging (when enabled)
3. Email notifications (when enabled)

Log files contain:
- Timestamp
- Log level
- Detailed message
- Error information (when applicable)

## Error Handling

The script includes comprehensive error handling:
- File system operations
- SFTP transfers
- Email notifications
- Environment configuration
- Process management

Failed transfers are:
1. Logged with detailed error messages
2. Moved to the failed directory
3. Included in email notifications (if enabled)

## Support

For issues with:
- FileZilla Pro CLI: Contact FileZilla support
- Script functionality: Open an issue in the repository
- SFTP connectivity: Verify your network and credentials

## License

None

## Contributing

None