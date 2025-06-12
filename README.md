# SCOM 2019 Bulk Maintenance Mode GUI Tool

A PowerShell GUI tool to manage Maintenance Mode for multiple servers in **SCOM 2019**. Supports over 500 servers in one go with real-time status updates.

## 📁 Script

**File Name:** `multiple_servers_mm_gui_2019.ps1`

## ✅ Features

- Load servers from a `.txt` file
- Start, Stop, or Update Maintenance Mode
- Predefined durations: 1hr, 2hr, 3hr, 4hr, 1 day, 1 week, 1 month
- Manual duration input (in minutes)
- Scrollable live status panel
- GUI-based success and failure messages
- Disables irrelevant inputs dynamically (e.g., duration for Stop)

## 📂 Folder Structure

📦SCOM-MM-Tool
- 📜 multiple_servers_mm_gui_2019.ps1
- 📄 servers.txt
- 📄 README.md



## ⚙️ Requirements

- PowerShell 5.1+
- .NET Framework 4.7 or higher
- SCOM 2019 Console & SDK installed
- Admin rights on SCOM

## 🚀 Usage

1. Add server names (one per line) to `servers.txt`
2. Run the script:
   ```powershell
   .\multiple_servers_mm_gui_2019.ps1



Made with ❤️ for SCOM admins.
