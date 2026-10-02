# Variables
$csvPath = "C:\Path\Machines_win11.csv"
$sourceFolder = "C:\Path"
$destinationFolder = "C$\Temp"
$logFile = "C:\Path\robocopy_log_Win11.txt"
$machines = Import-Csv -Path $csvPath

# Command
foreach ($entry in $machines) {
    $machine = $entry.Hostname
    $destinationPath = "\\$machine\$destinationFolder"
    $startTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Check if machine is online
    if (Test-Connection -ComputerName $machine -Count 2 -Quiet) {
        Add-Content -Path $logFile -Value "[${startTimestamp}] $machine is online. Starting copy..."

        try {
            # Run robocopy and capture output
            $output = robocopy $sourceFolder $destinationPath *.iso /Z /COPY:D /NDL /NP

            # Log robocopy output
            Add-Content -Path $logFile -Value $output

            if ($LASTEXITCODE -ge 8) {
                $errorMsg = "[${startTimestamp}] Robocopy failed on $machine with exit code $LASTEXITCODE"
                Add-Content -Path $logFile -Value $errorMsg
            } else {
                $endTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $logFile -Value "[${endTimestamp}] Successfully copied to $machine"
            }
        } catch {
            $errorMsg = "[${startTimestamp}] Exception while copying to $machine : $($_.Exception.Message)"
            Add-Content -Path $logFile -Value $errorMsg
        }
    } else {
        Add-Content -Path $logFile -Value "[${startTimestamp}] $machine is offline. Skipping..."
    }
}
