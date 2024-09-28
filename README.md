# PowerShell Projects

Welcome to the PowerShell Projects repository! This repository contains various PowerShell scripts designed to automate tasks, enhance productivity, and provide useful utilities for Windows users. 

## Contents

- [About This Repository](#about-this-repository)
- [Pinned Folders Backup and Restore Script](#pinned-folders-backup-and-restore-script)
  - [Features](#features)
  - [Usage](#usage)
  - [Setup](#setup)
  - [Running the Script](#running-the-script)
- [Contributing](#contributing)
- [License](#license)

## About This Repository

This repository hosts a collection of PowerShell scripts for various automation tasks. The primary script in this repository allows users to back up and restore pinned folders in Windows Quick Access, making it easy to save and reload pinned folder configurations.

## Pinned Folders Backup and Restore Script

This script allows you to back up or restore your pinned folders in Windows Quick Access. You can save the state of your pinned folders and restore them later, which is particularly useful when setting up a new machine or reconfiguring Quick Access.

### Features

- **Backup Mode:** Save the paths of currently pinned folders in Quick Access to a timestamped backup file.
- **Restore Mode:** Restore pinned folders from a backup file, automatically pinning them back to Quick Access.
- **Simple and interactive:** The script prompts the user for inputs, making it easy to use without prior experience.

### Usage

1. **Backup Mode**: The script will prompt you to enter a short name for the backup, which it will use along with the current timestamp to create a backup file.
2. **Restore Mode**: The script reads a specified backup file or automatically selects the most recent one, then restores the pinned folders listed in that file.

### Setup

1. Clone the repository or download the script directly from GitHub.

   ```bash
   git clone https://github.com/yourusername/powershell-projects.git
