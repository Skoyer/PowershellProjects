# Define backup directory
$backupDir = "C:\backups"

# Create backup directory if it doesn't exist
if (-not (Test-Path -Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir | Out-Null
}

# Get Quick Access folder
$folderPath = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}"
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace($folderPath)

if ($folder) {
    # Prompt for a short string for the backup name
    $backupName = Read-Host "Enter a short string for this backup"

    # Get current date and time for filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # Create backup filename
    $backupFileName = "$backupDir\$backupName-$timestamp.qlbak"

    # Create a new file for backup
    $backupFile = New-Item -Path $backupFileName -ItemType File

    foreach ($item in $folder.Items()) {
        if ($item.IsFolder) {
            # Try to get path using GetDetailsOf
            $folderPath = $folder.GetDetailsOf($item, 0)
            
            # If path is empty, try alternative methods
            if (-not $folderPath -or $folderPath -eq '') {
                # Try to get path from item's path property
                $folderPath = $item.Path
                if (-not $folderPath -or $folderPath -eq '') {
                    # If still empty, use the name as a fallback
                    $folderPath = $item.Name
                    Write-Output "Using item name as path for: $folderPath"
                } else {
                    Write-Output "Got path from item.Path for: $folderPath"
                }
            } else {
                Write-Output "Got path from GetDetailsOf for: $folderPath"
            }

            # Check if the path exists or is valid before adding to backup
            if (Test-Path -Path $folderPath -ErrorAction SilentlyContinue) {
                Add-Content -Path $backupFile.FullName -Value $folderPath
            } else {
                Write-Output "Skipping invalid path or non-existent path for: $folderPath"
            }
        }
    }
    Write-Output "Backup completed. File saved as: $backupFileName"
} else {
    Write-Output "Failed to access Quick Access folder."
}