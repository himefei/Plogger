# Load necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Logging Duration"
$form.Size = New-Object System.Drawing.Size(300, 180)
$form.StartPosition = "CenterScreen"

# Create buttons
$button60Sec = New-Object System.Windows.Forms.Button
$button60Sec.Text = "Logging 60 sec"
$button60Sec.Size = New-Object System.Drawing.Size(250, 30)
$button60Sec.Location = New-Object System.Drawing.Point(15, 10)

$button1Hour = New-Object System.Windows.Forms.Button
$button1Hour.Text = "Logging 1 hour"
$button1Hour.Size = New-Object System.Drawing.Size(250, 30)
$button1Hour.Location = New-Object System.Drawing.Point(15, 50)

$button3Hours = New-Object System.Windows.Forms.Button
$button3Hours.Text = "Logging 3 hours"
$button3Hours.Size = New-Object System.Drawing.Size(250, 30)
$button3Hours.Location = New-Object System.Drawing.Point(15, 90)

# Add buttons to the form
$form.Controls.Add($button60Sec)
$form.Controls.Add($button1Hour)
$form.Controls.Add($button3Hours)

# Variables to store the selected duration
$script:Interval = 1          # Interval in seconds
$script:TotalDuration = 0     # Total duration in seconds (set based on user selection)

# Button click event handlers
$button60Sec.Add_Click({
    $script:TotalDuration = 60
    $form.Close()
})

$button1Hour.Add_Click({
    $script:TotalDuration = 3600
    $form.Close()
})

$button3Hours.Add_Click({
    $script:TotalDuration = 10800
    $form.Close()
})

# Show the form
[void]$form.ShowDialog()

# If the user closes the form without selecting an option, exit the script
if ($TotalDuration -eq 0) {
    Write-Host "No option selected. Exiting the script."
    exit
}

# Proceed with the logging process
Write-Host "Logging CPU and Memory usage for $TotalDuration seconds..."

# Get the computer's serial number
$serialNumber = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber

# Replace any invalid filename characters in the serial number
$serialNumber = $serialNumber -replace '[\\\/\:\*\?"<>\|]', '_'

# Check if serial number is valid
if (-not $serialNumber -or $serialNumber -match '^(To be filled|Default|O\.E\.M\.)') {
    # Use the computer's name as an alternative
    $serialNumber = $env:COMPUTERNAME
}

# Generate a timestamp for the filename
$timestamp = Get-Date -Format "HHmmss_ddMMyyyy"  # Format: HHMMSS_DDMMYYYY

# Set the output file with serial number and timestamp in the filename
$OutputFile = Join-Path -Path (Get-Location).Path -ChildPath "_${serialNumber}_${timestamp}.csv"

# Remove the output file if it exists (unlikely with unique filenames)
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}

# Calculate the number of iterations
$Iterations = [Math]::Ceiling($TotalDuration / $Interval)

# Get the number of logical processors
$NumberOfLogicalProcessors = (Get-CimInstance -ClassName Win32_ComputerSystem).NumberOfLogicalProcessors

# Loop to capture data every interval
for ($i = 1; $i -le $Iterations; $i++) {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Get overall CPU usage
    try {
        $cpuTotal = Get-Counter '\Processor(_Total)\% Processor Time'
        $TotalCpuUsage = [math]::Round($cpuTotal.CounterSamples[0].CookedValue, 2)
    }
    catch {
        Write-Error "Error retrieving total CPU usage: $_"
        $TotalCpuUsage = 0
    }

    # Get overall memory usage using performance counters
    try {
        $memCounter = Get-Counter '\Memory\% Committed Bytes In Use'
        $MemoryUsagePercent = [math]::Round($memCounter.CounterSamples[0].CookedValue, 2)
    }
    catch {
        Write-Error "Error retrieving memory usage: $_"
        $MemoryUsagePercent = 0
    }

    # Get CPU usage data per process
    try {
        $Processes = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process |
            Where-Object { $_.Name -notin '_Total', 'Idle' } |
            Select-Object @{Name='Timestamp';Expression={$Timestamp}},
                           Name,
                           IDProcess,
                           @{Name='PercentProcessorTime';Expression={[math]::Round($_.PercentProcessorTime / $NumberOfLogicalProcessors, 2)}},
                           @{Name='MemoryUsagePercent';Expression={''}}  # Add MemoryUsagePercent as empty string
    }
    catch {
        Write-Error "Error retrieving process data: $_"
        continue
    }

    # Add the overall CPU and Memory usage as separate entries
    $TotalCpuEntry = [PSCustomObject]@{
        Timestamp = $Timestamp
        Name = '_TotalCPU'
        IDProcess = ''
        PercentProcessorTime = $TotalCpuUsage
        MemoryUsagePercent = ''
    }

    $TotalMemoryEntry = [PSCustomObject]@{
        Timestamp = $Timestamp
        Name = '_TotalMemory'
        IDProcess = ''
        PercentProcessorTime = ''
        MemoryUsagePercent = $MemoryUsagePercent
    }

    # Combine the total CPU and Memory usage entries with the processes
    $AllEntries = $Processes + $TotalCpuEntry + $TotalMemoryEntry

    # Export data to CSV
    try {
        if ($i -eq 1) {
            # Write headers on the first iteration
            $AllEntries | Export-Csv -Path $OutputFile -NoTypeInformation
        } else {
            # Append data without headers
            $AllEntries | Export-Csv -Path $OutputFile -NoTypeInformation -Append
        }
    }
    catch {
        Write-Error "Error writing to CSV file: $_"
    }

    # Update progress animation
    $ProgressPercent = ($i / $Iterations) * 100
    Write-Progress -Activity "Logging CPU and Memory Usage" -Status ("Progress: {0:N0}% Complete" -f $ProgressPercent) -PercentComplete $ProgressPercent

    # Wait for the specified interval
    Start-Sleep -Seconds $Interval
}

# Clear the progress bar
Write-Progress -Activity "Logging CPU and Memory Usage" -Completed

# Open the folder containing the CSV file
$OutputFolder = Split-Path -Path $OutputFile -Parent
Start-Process -FilePath $OutputFolder
