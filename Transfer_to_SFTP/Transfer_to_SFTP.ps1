# Requires System.Web assembly for HTML encoding
Add-Type -AssemblyName System.Web

# Temporary file for capturing output
$outputFile = Join-Path $env:TEMP "script_output.txt"

Function Load-EnvVariables {
    param(
        [string]$EnvFilePath = ".\.env"
    )
    
    if (!(Test-Path $EnvFilePath)) {
        Write-Host "No .env file found at $EnvFilePath. Exiting."
        exit 1
    }

    $lines = Get-Content $EnvFilePath
    
    foreach ($line in $lines) {
        if ($line -match "^\s*$" -or $line -match "^\s*#") {
            continue
        }
        
        $parts = $line -split "=", 2
        if ($parts.Count -ne 2) {
            continue
        }
        
        $key = $parts[0].Trim()
        $value = $parts[1].Trim()
        
        # Remove surrounding quotes if present
        if ($value.StartsWith("'") -and $value.EndsWith("'")) {
            $value = $value.Trim("'")
        } elseif ($value.StartsWith('"') -and $value.EndsWith('"')) {
            $value = $value.Trim('"')
        }
        
        # Special handling for path variables
        if ($key -like "*PATH*") {
            if ($key -eq "SFTP_UPLOAD_PATH") {
                # For SFTP paths, ensure forward slashes and no trailing slash
                $value = $value.Replace('\', '/').TrimEnd('/')
            } else {
                # For local Windows paths, ensure one trailing backslash
                $value = $value.TrimEnd('\') + '\'
                # Convert to absolute path if relative
                if (![System.IO.Path]::IsPathRooted($value)) {
                    $value = Join-Path $PWD $value
                }
            }
        }
        
        Set-Item -Path "env:$key" -Value $value
        
        # Log loaded variables except passwords
        if ($env:DEBUG -eq "True") {
            if ($key -like "*PASSWORD*") {
                Write-Host "Loaded $key = ********"
            } else {
                Write-Host "Loaded $key = $value"
            }
        }
    }
}

Function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $debugEnabled = $env:DEBUG -eq "True"
    $logToFile = $env:LOG_TO_FILE -eq "True"
    $timeStamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $logMessage = "[$timeStamp] [$Level] $Message"

    if ($Level -ne "DEBUG" -or ($Level -eq "DEBUG" -and $debugEnabled)) {
        # Write to both console and output file without extra newlines
        $logMessage | Write-Host
        $logMessage | Out-File -FilePath $outputFile -Append -NoNewline
        "" | Out-File -FilePath $outputFile -Append # Single newline after message
        
        if ($logToFile) {
            $logDir = $env:LOG_PATH
            if (!(Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            $logFile = Join-Path $logDir $env:LOG_FILE
            $logMessage | Out-File -FilePath $logFile -Append
        }
    }
}

Function Send-Email {
    param(
        [string]$Subject,
        [string]$Body
    )
    if ($env:SEND_EMAIL -eq "True") {
        Write-Log "Sending email to $($env:EMAIL_TO)" "INFO"
        try {
            # Create Mail Message
            $mailMessage = [System.Net.Mail.MailMessage]::new()
            $mailMessage.From = [System.Net.Mail.MailAddress]::new($env:EMAIL_FROM)
            $mailMessage.To.Add($env:EMAIL_TO)
            $mailMessage.Subject = $Subject
            $mailMessage.Body = $Body
            $mailMessage.IsBodyHtml = $true
            $mailMessage.Priority = [System.Net.Mail.MailPriority]::High

            # Create SMTP Client
            $smtpClient = [System.Net.Mail.SmtpClient]::new()
            $smtpClient.Host = $env:SMTP_HOST
            $smtpClient.Port = [int]$env:SMTP_PORT

            # Set credentials
            $securePassword = $env:SMTP_PASSWORD | ConvertTo-SecureString -AsPlainText -Force
            $credentials = [System.Net.NetworkCredential]::new($env:SMTP_USERNAME, $securePassword)
            $smtpClient.Credentials = $credentials

            # Set SSL/TLS settings
            if ($env:SMTP_TTLS -eq "True") {
                $smtpClient.EnableSsl = $true
            }

            # Send the email
            $smtpClient.Send($mailMessage)
            Write-Log "Email sent successfully." "INFO"

            # Clean up
            $mailMessage.Dispose()
            $smtpClient.Dispose()
        }
        catch {
            Write-Log "Error sending email: $($_)" "ERROR"
        }
    }
}

Function Upload-FileWithFileZilla {
    param(
        [string]$LocalFile
    )
    
    if ($env:UPLOAD_TO_SFTP -ne "True") {
        Write-Log "UPLOAD_TO_SFTP is not enabled. Skipping upload." "DEBUG"
        return $true
    }

    Write-Log "Preparing to upload file: $($LocalFile)" "INFO"
    
    try {
        # Validate FileZilla CLI path (remove trailing backslash)
        $filezillaPath = $env:FILEZILLA_PATH.TrimEnd('\')
        
        # Verify FileZilla CLI exists
        if (!(Test-Path $filezillaPath)) {
            Write-Log "FileZilla CLI not found at $filezillaPath" "ERROR"
            return $false
        }
        
        # Escape the local file path and prepare remote path
        $escapedLocalFile = $LocalFile.Replace('\', '\\')
        $fileName = Split-Path $LocalFile -Leaf
        
        # Use SFTP_UPLOAD_PATH, ensuring it starts with a forward slash and doesn't end with one
        $remotePath = $env:SFTP_UPLOAD_PATH
        if (!$remotePath.StartsWith('/')) {
            $remotePath = '/' + $remotePath
        }
        $remotePath = $remotePath.TrimEnd('/')
        
        # Construct full remote file path
        $remoteFilePath = "$remotePath/$fileName"
        
        # Create temporary script file
        $scriptPath = Join-Path $env:TEMP "fz_script_$(Get-Random).txt"
        
        # Create the script content with fully escaped path and correct put command
        $scriptContent = @"
# Connection settings
connect --user $($env:SFTP_USERNAME) --pass $($env:SFTP_PASSWORD) sftp://$($env:SFTP_SERVER_HOST):$($env:SFTP_SERVER_PORT)

# Upload file with overwrite
put --exists overwrite "$escapedLocalFile" "$remoteFilePath"
ls "$remotePath"
exit
"@
        
        # Write script content to file
        $scriptContent | Out-File -FilePath $scriptPath -Encoding ASCII
        
        Write-Log "Remote upload details:" "INFO"
        Write-Log "  Server: $($env:SFTP_SERVER_HOST):$($env:SFTP_SERVER_PORT)" "INFO"
        Write-Log "  Username: $($env:SFTP_USERNAME)" "INFO"
        Write-Log "  Remote Path: $remoteFilePath" "INFO"
        
        # Write script content to file
        $scriptContent | Out-File -FilePath $scriptPath -Encoding ASCII
        
        # Execute FileZilla CLI with script
        $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processStartInfo.FileName = $filezillaPath
        $processStartInfo.Arguments = "--script `"$scriptPath`""
        $processStartInfo.UseShellExecute = $false
        $processStartInfo.RedirectStandardOutput = $true
        $processStartInfo.RedirectStandardError = $true
        $processStartInfo.CreateNoWindow = $true  # Hide window since we're providing password
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processStartInfo
        
        $outputBuilder = New-Object System.Text.StringBuilder
        $errorBuilder = New-Object System.Text.StringBuilder
        
        $outputHandler = {
            if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
                $outputBuilder.AppendLine($EventArgs.Data)
                Write-Log "FZ Output: $($EventArgs.Data)" "DEBUG"
            }
        }
        
        $errorHandler = {
            if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
                $errorBuilder.AppendLine($EventArgs.Data)
                Write-Log "FZ Error: $($EventArgs.Data)" "ERROR"
            }
        }
        
        # Register the event handlers
        $outputEvent = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action $outputHandler
        $errorEvent = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action $errorHandler
        
        Write-Log "Starting FileZilla process..." "DEBUG"
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        
        # Wait for process with timeout
        $processExited = $process.WaitForExit(120000) # 2 minute timeout
        
        # Clean up event handlers first
        if ($outputEvent) {
            Unregister-Event -SourceIdentifier $outputEvent.Name
            Remove-Job -Id $outputEvent.Id -Force
        }
        if ($errorEvent) {
            Unregister-Event -SourceIdentifier $errorEvent.Name
            Remove-Job -Id $errorEvent.Id -Force
        }
        
        # Kill the process if it hasn't exited
        if (!$processExited) {
            Write-Log "FileZilla process did not exit within timeout period" "ERROR"
            try {
                $process.Kill()
                $process.WaitForExit(5000) # Give it 5 seconds to die
            } catch {
                Write-Log "Could not terminate FileZilla process: $_" "ERROR"
            }
        }
        
        # Now that process is definitely done, clean up the script file
        Start-Sleep -Seconds 1 # Small delay to ensure file handle is released
        try {
            Remove-Item $scriptPath -Force -ErrorAction Stop
            Write-Log "Removed temporary script file" "DEBUG"
        } catch {
            Write-Log "Could not remove temporary script file: $_" "DEBUG"
        }
        
        $output = $outputBuilder.ToString()
        $error = $errorBuilder.ToString()
        
        if ($processExited -and $process.ExitCode -eq 0) {
            Write-Log "FileZilla CLI successfully uploaded file" "INFO"
            Write-Log "Local file: $LocalFile" "INFO"
            Write-Log "Remote location: $remoteFilePath" "INFO"
            
            # Parse and log the directory listing if available
            if ($output -match "ls.*\n(.*)\n") {
                Write-Log "Remote directory contents after upload:" "INFO"
                $dirListing = $matches[1] -split '\n' | ForEach-Object { "  $_" }
                $dirListing | ForEach-Object { Write-Log $_ "INFO" }
            }
            
            return $true
        } else {
            Write-Log "FileZilla CLI failed to upload $($LocalFile)" "ERROR"
            Write-Log "Local file: $LocalFile" "ERROR"
            Write-Log "Intended remote location: $remoteFilePath" "ERROR"
            Write-Log "Last command output: $output" "ERROR"
            Write-Log "Error output: $error" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error running FileZilla CLI for $($LocalFile) - $($_)" "ERROR"
        return $false
    }
}

Function Process-Files {
    Write-Log "Scanning for files in $($env:FILE_PATH)" "INFO"
    
    # Check if directories exist
    $requiredPaths = @($env:FILE_PATH, $env:FILE_PATH_PROCESSED, $env:FILE_PATH_FAILED)
    foreach ($path in $requiredPaths) {
        if (!(Test-Path $path)) {
            Write-Log "Creating directory: $path" "INFO"
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }

    $sourceFiles = Get-ChildItem -Path $env:FILE_PATH -File -Force -ErrorAction SilentlyContinue
    
    if ($null -eq $sourceFiles -or @($sourceFiles).Count -eq 0) {
        Write-Log "No files found in $($env:FILE_PATH)" "INFO"
        return
    }

    foreach ($file in $sourceFiles) {
        Write-Log "Processing file: $($file.FullName)" "INFO"
        $uploadSuccess = Upload-FileWithFileZilla -LocalFile $file.FullName
        
        $destinationFolder = if ($uploadSuccess) { $env:FILE_PATH_PROCESSED } else { $env:FILE_PATH_FAILED }
        $destination = Join-Path $destinationFolder $file.Name
        
        try {
            Move-Item -Path $file.FullName -Destination $destination -Force
            Write-Log "Moved $($file.FullName) to $($destination)" "INFO"
        } catch {
            Write-Log "Failed to move $($file.FullName) to $($destination) - $($_)" "ERROR"
        }
    }
}

# Clean up any existing output file
if (Test-Path $outputFile) {
    Remove-Item $outputFile -Force
}

# Main script execution
Load-EnvVariables
Write-Log "Starting file transfer script" "INFO"
Process-Files

if ($env:SEND_EMAIL -eq "True") {
    $subject = $env:EMAIL_SUBJECT
    
    # Read the output file
    $logContent = ""
    if (Test-Path $outputFile) {
        $logContent = Get-Content -Path $outputFile -Raw
        # Remove any extra blank lines
        $logContent = $logContent -replace "`r`n`r`n", "`r`n"
        # Escape any HTML characters and convert newlines to <br>
        $logContent = [System.Web.HttpUtility]::HtmlEncode($logContent)
        $logContent = $logContent.Replace("`r`n", "<br>")
    }
    
    $body = @"
<html>
<body style='font-family: Arial, sans-serif;'>
<p>File transfer process completed at $(Get-Date).</p>
<h3>Script Output:</h3>
<pre style='background-color: #f4f4f4; padding: 15px; border-radius: 5px; font-family: Consolas, monospace; line-height: 1.4;'>
$logContent
</pre>
</body>
</html>
"@
    
    Send-Email -Subject $subject -Body $body
}

Write-Log "Script execution finished." "INFO"

# Clean up the output file
if (Test-Path $outputFile) {
    Remove-Item $outputFile -Force
}