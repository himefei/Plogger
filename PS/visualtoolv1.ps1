# Load necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create an OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$openFileDialog.Filter = "CSV files (*.csv)|*.csv"
$openFileDialog.Title = "Select the CSV File to Generate the Report"

# Show the dialog and get the result
$dialogResult = $openFileDialog.ShowDialog()

# Check if the user selected a file or canceled
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $CsvFile = $openFileDialog.FileName
    Write-Host "Selected CSV file: $CsvFile"
} else {
    Write-Host "No file selected. Exiting the script."
    exit
}

# Specify the HTML report file name to match the CSV file name but with .html extension
$HtmlReportFile = [System.IO.Path]::ChangeExtension($CsvFile, '.html')

# Check if the CSV file exists
if (-Not (Test-Path $CsvFile)) {
    Write-Error "CSV file '$CsvFile' not found."
    exit
}

# Read the CSV file
$Data = Import-Csv -Path $CsvFile

# Check if data was imported successfully
if (-Not $Data) {
    Write-Error "No data found in CSV file."
    exit
}

# Ensure all entries have consistent properties
$Data = $Data | ForEach-Object {
    [PSCustomObject]@{
        Timestamp = $_.Timestamp
        Name = $_.Name
        IDProcess = $_.IDProcess
        PercentProcessorTime = if ($_.PercentProcessorTime -ne '') {
            try {
                [float]::Parse($_.PercentProcessorTime, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                Write-Warning "Failed to parse PercentProcessorTime '$($_.PercentProcessorTime)' at Timestamp '$($_.Timestamp)'. Defaulting to 0."
                0
            }
        } else {
            0
        }
        MemoryUsagePercent = if ($_.MemoryUsagePercent -ne '') {
            try {
                [float]::Parse($_.MemoryUsagePercent, [System.Globalization.CultureInfo]::InvariantCulture)
            } catch {
                Write-Warning "Failed to parse MemoryUsagePercent '$($_.MemoryUsagePercent)' at Timestamp '$($_.Timestamp)'. Defaulting to 0."
                0
            }
        } else {
            0
        }
    }
}

# Exclude "Idle" processes, but include '_TotalCPU' and '_TotalMemory'
$excludedProcesses = @('Idle', 'System Idle Process')
$Data = $Data | Where-Object {
    ($excludedProcesses -notcontains $_.Name -and -not ($_.Name -match 'idle')) -or $_.Name -eq '_TotalCPU' -or $_.Name -eq '_TotalMemory'
}

# Extract overall CPU usage data (where Name is '_TotalCPU')
$TotalCpuData = $Data | Where-Object { $_.Name -eq '_TotalCPU' } | Sort-Object Timestamp
$TotalCpuTimestamps = $TotalCpuData | Select-Object -ExpandProperty Timestamp
$TotalCpuUsages = $TotalCpuData | Select-Object -ExpandProperty PercentProcessorTime

# Extract overall Memory usage data (where Name is '_TotalMemory')
$TotalMemoryData = $Data | Where-Object { $_.Name -eq '_TotalMemory' } | Sort-Object Timestamp
$TotalMemoryTimestamps = $TotalMemoryData | Select-Object -ExpandProperty Timestamp
$TotalMemoryUsages = $TotalMemoryData | Select-Object -ExpandProperty MemoryUsagePercent

# Remove '_TotalCPU' and '_TotalMemory' from the main data to avoid duplication
$Data = $Data | Where-Object { $_.Name -ne '_TotalCPU' -and $_.Name -ne '_TotalMemory' }

# Exclude 'PercentProcessorTime' with 0 (since we assigned 0 for non-process entries)
$Data = $Data | Where-Object { $_.PercentProcessorTime -ne 0 }

# Convert PercentProcessorTime to float and add BaseName
$Data = $Data | ForEach-Object {
    # Extract base process name by removing any instance numbers (e.g., #1, #2)
    $BaseName = $_.Name -replace '#\d+$', ''
    # Create a new object with all existing properties plus BaseName
    $_ | Select-Object *, @{Name='BaseName';Expression={$BaseName}}
}

# Get all unique timestamps
$AllTimestamps = $Data | Select-Object -ExpandProperty Timestamp | Sort-Object -Unique

# Get unique base process names
$BaseProcessNames = $Data | Select-Object -ExpandProperty BaseName | Sort-Object -Unique

# Group data by BaseName and Timestamp and sum CPU usage
$GroupedData = $Data | Group-Object -Property BaseName, Timestamp | ForEach-Object {
    $BaseName = $_.Group[0].BaseName
    $Timestamp = $_.Group[0].Timestamp
    $TotalCpuUsage = ($_.Group | Measure-Object -Property PercentProcessorTime -Sum).Sum

    # Create a custom object with BaseName, Timestamp, TotalCpuUsage
    [PSCustomObject]@{
        BaseName = $BaseName
        Timestamp = $Timestamp
        TotalCpuUsage = $TotalCpuUsage
    }
}

# Build a lookup table for quick access
$DataLookup = @{}
foreach ($Entry in $GroupedData) {
    $BaseName = $Entry.BaseName
    $Timestamp = $Entry.Timestamp
    $TotalCpuUsage = $Entry.TotalCpuUsage

    if (-not $DataLookup.ContainsKey($BaseName)) {
        $DataLookup[$BaseName] = @{}
    }
    $DataLookup[$BaseName][$Timestamp] = $TotalCpuUsage
}

# Prepare data for JavaScript and create ProcessStats
$ProcessData = @{}
$ProcessStats = @()

foreach ($BaseName in $BaseProcessNames) {
    $ProcessData[$BaseName] = @{
        'timestamps' = $AllTimestamps
        'cpuUsages' = @()
        'maxCpuUsage' = 0
    }

    foreach ($Timestamp in $AllTimestamps) {
        $TotalCpuUsage = 0
        if ($DataLookup[$BaseName] -and $DataLookup[$BaseName].ContainsKey($Timestamp)) {
            $TotalCpuUsage = $DataLookup[$BaseName][$Timestamp]
        }
        $ProcessData[$BaseName]['cpuUsages'] += $TotalCpuUsage

        # Update maxCpuUsage
        if ($TotalCpuUsage -gt $ProcessData[$BaseName]['maxCpuUsage']) {
            $ProcessData[$BaseName]['maxCpuUsage'] = $TotalCpuUsage
        }
    }

    # Round maxCpuUsage to two decimal places
    $RoundedMaxCpuUsage = [math]::Round($ProcessData[$BaseName]['maxCpuUsage'], 2)

    # Add to ProcessStats for sorting
    $ProcessStats += [PSCustomObject]@{
        Name = $BaseName
        MaxCpuUsage = $RoundedMaxCpuUsage
    }

    # Update the maxCpuUsage in ProcessData as well
    $ProcessData[$BaseName]['maxCpuUsage'] = $RoundedMaxCpuUsage
}

# Sort processes by maximum CPU usage (high to low)
$ProcessStats = $ProcessStats | Sort-Object -Property MaxCpuUsage -Descending

# Find the global time range
$AllTimestampsAllData = $TotalCpuTimestamps + $AllTimestamps + $TotalMemoryTimestamps
$GlobalMinTime = [DateTime]::ParseExact(($AllTimestampsAllData | Sort-Object)[0], 'yyyy-MM-dd HH:mm:ss', $null)
$GlobalMaxTime = [DateTime]::ParseExact(($AllTimestampsAllData | Sort-Object)[-1], 'yyyy-MM-dd HH:mm:ss', $null)

# Convert times to ISO 8601 format for JavaScript
$GlobalMinTimeISO = $GlobalMinTime.ToString("yyyy-MM-ddTHH:mm:ss")
$GlobalMaxTimeISO = $GlobalMaxTime.ToString("yyyy-MM-ddTHH:mm:ss")

# Convert the data to JSON format for JavaScript
$JsonProcessData = ConvertTo-Json -InputObject $ProcessData -Depth 6

# Convert overall CPU usage data to JSON
$TotalCpuDataJson = ConvertTo-Json -InputObject @{
    'timestamps' = $TotalCpuTimestamps
    'cpuUsages' = $TotalCpuUsages
} -Depth 5

# Convert overall Memory usage data to JSON
$TotalMemoryDataJson = ConvertTo-Json -InputObject @{
    'timestamps' = $TotalMemoryTimestamps
    'memoryUsages' = $TotalMemoryUsages
} -Depth 5

# Generate the HTML options for the dropdown using sorted ProcessStats
$OptionTags = ""
foreach ($ProcessStat in $ProcessStats) {
    $BaseName = $ProcessStat.Name
    $MaxCpuUsage = $ProcessStat.MaxCpuUsage

    # Format MaxCpuUsage to show two decimal places
    $MaxCpuUsageFormatted = "{0:F2}" -f $MaxCpuUsage

    $ProcessEscaped = [System.Web.HttpUtility]::HtmlEncode($BaseName)
    # Determine color indicator based on MaxCpuUsage
    if ($MaxCpuUsage -gt 20) {
        $ColorIndicator = "🔴" # Red circle
    } elseif ($MaxCpuUsage -ge 5 -and $MaxCpuUsage -le 20) {
        $ColorIndicator = "🟡" # Yellow circle
    } else {
        $ColorIndicator = "🟢" # Green circle
    }
    $OptionTags += "        <option value='$ProcessEscaped'>$ColorIndicator $ProcessEscaped (Max CPU: $MaxCpuUsageFormatted%)</option>`n"
}

# Generate the HTML content
$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Process CPU and Memory Usage Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }

        #combinedChartContainer, #processChartContainer {
            width: 100%;
            height: 500px;
        }

        #processSelect {
            padding: 5px;
            font-size: 14px;
        }

        .chart-section {
            margin-bottom: 50px;
        }
    </style>
    <!-- Include Chart.js library from CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Include Luxon for date parsing (used by Chart.js) -->
    <script src="https://cdn.jsdelivr.net/npm/luxon@1/build/global/luxon.min.js"></script>
    <!-- Include Chart.js adapters for Luxon -->
    <script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-luxon"></script>
</head>
<body>
    <h1>Process CPU and Memory Usage Report</h1>

    <div id="combinedChartContainer" class="chart-section">
        <h2>Overall CPU and Memory Usage (%) Over Time</h2>
        <canvas id="combinedChart"></canvas>
    </div>

    <label for="processSelect"><strong>Select a Process:</strong></label>
    <select id="processSelect">
        <option value="" disabled selected>Select a process</option>
$OptionTags
    </select>

    <div id="processChartContainer" class="chart-section">
        <h2>Process CPU Usage (%) Over Time</h2>
        <canvas id="cpuChart"></canvas>
    </div>

    <script>
        // Process data from PowerShell
        var processData = $($JsonProcessData);
        var totalCpuData = $($TotalCpuDataJson);
        var totalMemoryData = $($TotalMemoryDataJson);

        // Global time range
        var globalMinTime = new Date('$GlobalMinTimeISO');
        var globalMaxTime = new Date('$GlobalMaxTimeISO');

        // Elements
        var processSelect = document.getElementById('processSelect');
        var cpuChartCanvas = document.getElementById('cpuChart');
        var combinedChartCanvas = document.getElementById('combinedChart');
        var cpuChart;
        var combinedChart;

        // Create combined CPU and Memory usage chart
        var combinedTimestamps = totalCpuData.timestamps;
        var cpuUsages = totalCpuData.cpuUsages.map(function(val) { return parseFloat(val); });
        var memoryUsages = totalMemoryData.memoryUsages.map(function(val) { return parseFloat(val); });

        combinedChart = new Chart(combinedChartCanvas, {
            type: 'line',
            data: {
                labels: combinedTimestamps,
                datasets: [
                    {
                        label: 'CPU Usage (%)',
                        data: cpuUsages,
                        borderColor: 'rgba(255, 99, 132, 1)', // Red color
                        backgroundColor: 'rgba(255, 99, 132, 0.1)',
                        borderWidth: 2,
                        fill: false,
                        pointRadius: 1,
                        tension: 0.1
                    },
                    {
                        label: 'Memory Usage (%)',
                        data: memoryUsages,
                        borderColor: 'rgba(75, 192, 192, 1)', // Green color
                        backgroundColor: 'rgba(75, 192, 192, 0.1)',
                        borderWidth: 2,
                        fill: false,
                        pointRadius: 1,
                        tension: 0.1
                    }
                ]
            },
            options: {
                scales: {
                    x: {
                        type: 'time',
                        time: {
                            parser: 'yyyy-MM-dd HH:mm:ss',
                            tooltipFormat: 'HH:mm:ss',
                            unit: 'minute',
                            displayFormats: {
                                minute: 'HH:mm:ss'
                            }
                        },
                        min: globalMinTime,
                        max: globalMaxTime,
                        title: {
                            display: true,
                            text: 'Time'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        max: 100,
                        title: {
                            display: true,
                            text: 'Usage (%)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top'
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false
                    }
                },
                responsive: true,
                maintainAspectRatio: false
            }
        });

        // Event listener for process selection
        processSelect.addEventListener('change', function() {
            var selectedProcess = this.value;
            var data = processData[selectedProcess];
            if (data) {
                var timestamps = data.timestamps;
                var cpuUsages = data.cpuUsages.map(function(val) { return parseFloat(val); });

                // Destroy existing chart if it exists
                if (cpuChart) {
                    cpuChart.destroy();
                }

                // Create new chart
                cpuChart = new Chart(cpuChartCanvas, {
                    type: 'line',
                    data: {
                        labels: timestamps,
                        datasets: [{
                            label: 'CPU Usage (%)',
                            data: cpuUsages,
                            borderColor: 'rgba(54, 162, 235, 1)',
                            borderWidth: 2,
                            fill: false,
                            pointRadius: 1,
                            tension: 0.1
                        }]
                    },
                    options: {
                        scales: {
                            x: {
                                type: 'time',
                                time: {
                                    parser: 'yyyy-MM-dd HH:mm:ss',
                                    tooltipFormat: 'HH:mm:ss',
                                    unit: 'minute',
                                    displayFormats: {
                                        minute: 'HH:mm:ss'
                                    }
                                },
                                min: globalMinTime,
                                max: globalMaxTime,
                                title: {
                                    display: true,
                                    text: 'Time'
                                }
                            },
                            y: {
                                beginAtZero: true,
                                max: 100,
                                title: {
                                    display: true,
                                    text: 'CPU Usage (%)'
                                }
                            }
                        },
                        plugins: {
                            legend: {
                                display: false
                            },
                            tooltip: {
                                mode: 'index',
                                intersect: false
                            }
                        },
                        responsive: true,
                        maintainAspectRatio: false
                    }
                });
            }
        });
    </script>
</body>
</html>
"@

# Save the HTML content to a file
Set-Content -Path $HtmlReportFile -Value $HtmlContent -Encoding UTF8

# Open the HTML report in the default web browser
Start-Process -FilePath $HtmlReportFile

Write-Host "HTML report generated and opened: $HtmlReportFile"
