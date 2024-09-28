# Define backup directory
$backupDir = "C:\backups"

# Create backup directory if it doesn't exist
if (-not (Test-Path -Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir | Out-Null
}

# Prompt the user to choose between backup or restore
$operation = Read-Host "Enter 'b' to backup pinned folders or 'r' to restore pinned folders"

if ($operation -eq 'b') {
    # Backup operation
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
                $itemPath = $folder.GetDetailsOf($item, 0)
                
                # If path is empty, try alternative methods
                if (-not $itemPath -or $itemPath -eq '') {
                    # Try to get path from item's path property
                    $itemPath = $item.Path
                    if (-not $itemPath -or $itemPath -eq '') {
                        # If still empty, use the name as a fallback
                        $itemPath = $item.Name
                        Write-Output "Using item name as path for: $itemPath"
                    } else {
                        Write-Output "Got path from item.Path for: $itemPath"
                    }
                } else {
                    Write-Output "Got path from GetDetailsOf for: $itemPath"
                }

                # Check if the path exists or is valid before adding to backup
                if (Test-Path -Path $itemPath -ErrorAction SilentlyContinue) {
                    Add-Content -Path $backupFile.FullName -Value $itemPath
                } else {
                    Write-Output "Skipping invalid or non-existent path: $itemPath"
                }
            }
        }
        Write-Output "Backup completed. File saved as: $backupFileName"
    } else {
        Write-Output "Failed to access Quick Access folder."
    }

} elseif ($operation -eq 'r') {
    # Restore operation

    # Function to get the most recent backup file
    function Get-MostRecentBackup {
        Get-ChildItem -Path $backupDir -Filter "*.qlbak" | 
        Sort-Object LastWriteTime -Descending | 
        Select-Object -First 1
    }

    # Check if a specific backup file was provided as an argument
    if ($args.Count -gt 0) {
        # Check if the argument is a full path or just a filename
        if (Test-Path -Path $args[0]) {
            $backupFile = $args[0] # Full path provided
        } else {
            $backupFile = Join-Path -Path $backupDir -ChildPath "$($args[0]).qlbak"
        }
    } else {
        $backupFile = (Get-MostRecentBackup).FullName
    }

    # Check if the backup file exists
    if (-not (Test-Path -Path $backupFile)) {
        Write-Output "Backup file not found: $backupFile"
        exit
    }

    # Read content from the backup file
    $backupContent = Get-Content -Path $backupFile

    # Iterate through each line (each line represents a folder path)
    foreach ($path in $backupContent) {
        # Trim any whitespace and check if path is not empty
        if ($path.Trim() -ne '') {
            if (Test-Path -Path $path) {
                # Pin the folder to Quick Access
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace($path)
                if ($folder) {
                    $folder.Self.InvokeVerb("pintohome")
                }
            } else {
                Write-Output "Path not found: $path"
            }
        } else {
            Write-Output "Skipping empty path."
        }
    }

    Write-Output "Restoration from $backupFile completed."

} else {
    Write-Output "Invalid input. Please enter 'b' for backup or 'r' for restore."
}
