# Define backup directory
$backupDir = "C:\backups"

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
