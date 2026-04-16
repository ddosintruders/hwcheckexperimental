
# hwcheck

# 🔍 Laptop Inspector — Portable Edition

> **All-in-one Windows laptop diagnostic tool** — Run from a flash drive, no installation required.

[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D6?logo=windows)](https://www.microsoft.com/windows)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://docs.microsoft.com/powershell/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Portable](https://img.shields.io/badge/Portable-USB%20Ready-orange)](#usage)

---

## ✨ What It Does

Laptop Inspector is a **single-file**, portable Windows diagnostic tool that performs a comprehensive hardware and software audit of any laptop. Just double-click `LaptopInspector.bat` on any Windows machine and get a full inspection report in seconds.

Perfect for:
- 🛒 **Buying a used laptop** — verify specs before you pay
- 🔧 **IT technicians** — quickly audit machines in the field
- 📊 **Fleet management** — track laptop health over time with CSV history
- 🏫 **Schools & offices** — inspect donated or returned equipment

---

## 🚀 Features

### 🖥️ System Information
- CPU model, cores, and threads
- RAM capacity
- GPU and driver version
- Display resolution and refresh rate
- Manufacturer, model, and serial number

### 🔋 Battery Health (Deep Analysis)
- Current charge level
- Battery chemistry (Li-ion, LiPo, etc.)
- Design vs. full charge capacity
- **Wear level** percentage
- Cycle count
- Power plan detection

### 💾 Storage Diagnostics
- Physical disk detection (SSD/HDD)
- S.M.A.R.T. status monitoring
- Per-drive capacity and free space
- Usage percentage alerts

### 🛡️ Security & OS
- Windows activation status
- Antivirus (Defender) status
- Firewall profile detection
- BitLocker encryption status
- TPM version check
- OS build and install date

### 🌐 Network
- Wi-Fi adapter and signal strength
- Ethernet adapter detection
- Internet connectivity test with latency

### 🎛️ Peripherals
- Webcam detection
- Bluetooth adapter
- Audio devices
- USB device enumeration

### 📈 Performance
- Running process count
- Top 5 RAM consumers
- Startup program audit
- Critical system events (last 48 hours)

### 📝 Reports (Auto-generated)
| Format | Description |
|--------|-------------|
| **TXT** | Full text report for archival |
| **CSV** | Append-only history for tracking multiple inspections |
| **HTML** | Beautiful, styled report that opens in any browser |

---

## 🎯 Weighted Scoring System

Each check is weighted by importance:

| Weight | Checks |
|--------|--------|
| **x3** | Disk S.M.A.R.T. status |
| **x2** | CPU match, RAM, Battery level, Battery wear, Windows activation, System errors |
| **x1** | GPU, Resolution, Wi-Fi, Internet, Antivirus, Firewall, TPM, Webcam, Bluetooth, Audio, Storage capacity |

**Final Result:**
- 🟢 **PASS** — Score ≥ 80%
- 🟡 **WARNING** — Score 60–79%
- 🔴 **FAIL** — Score < 60%

---

## 📦 Usage

### Option 1: Run from USB Flash Drive
1. Copy `LaptopInspector.bat` to any USB flash drive
2. Plug the USB into the target laptop
3. Double-click `LaptopInspector.bat`
4. Click **START SCAN** in the GUI
5. Reports are saved to a `Reports/` folder next to the script

### Option 2: Run Directly
1. Download `LaptopInspector.bat`
2. Right-click → **Run as Administrator** (recommended for full access)
3. Click **START SCAN**

### Command Line (Optional)
```powershell
# Quick scan (skips slow checks like event log scanning)
powershell -ExecutionPolicy Bypass -File LaptopInspector.bat -Quick

# Full scan (default)
powershell -ExecutionPolicy Bypass -File LaptopInspector.bat -Full
```

> **Note:** The script auto-elevates via PowerShell. Some checks (temperature, BitLocker, TPM) require Administrator privileges for full results.

---

## 🖼️ GUI Interface

The tool features a modern dark-themed GUI built with WPF:

- **Left panel** — System dashboard with key specs at a glance
- **Right panel** — Real-time scrolling log with color-coded output
- **Progress bar** — Tracks scan completion (10 phases)
- **Score display** — Large result score with color coding
- **Report button** — One-click access to the HTML report

---

## 📁 Project Structure

```
LaptopInspector/
├── LaptopInspector.bat    # The main script (self-contained)
├── README.md              # This file
├── LICENSE                # MIT License
├── .gitignore             # Ignores generated reports
└── Reports/               # Auto-created on first run
    ├── report_YYYY-MM-DD_HH-mm-ss.txt
    ├── report_YYYY-MM-DD_HH-mm-ss.html
    └── history.csv
```

---

## ⚙️ Customization

Edit the `$expected` block at the top of the script to set your own target specs:

```powershell
$expected = @{
    CPU = "i7"                    # CPU must contain this string
    RAM = 8                       # Minimum RAM in GB
    GPU = "Intel"                 # GPU must contain this string
    BATTERY = 40                  # Minimum battery percentage
    STORAGE_MIN_GB = 200          # Minimum total storage in GB
    RESOLUTION_MIN_WIDTH = 1920   # Minimum horizontal resolution
}
```

---

## 🔒 Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 / 11 |
| **PowerShell** | 5.1+ (pre-installed on Windows 10+) |
| **Privileges** | Standard user (Admin recommended for full results) |
| **Dependencies** | None — fully self-contained |
| **Disk Space** | ~53 KB (single file) |

---

## 🤝 Contributing

Contributions are welcome! Here are some ideas:

- [ ] Add BIOS/UEFI version detection
- [ ] Add RAM slot details (speed, type, slots used)
- [ ] Add display panel info (IPS/TN, color depth)
- [ ] Add keyboard backlight detection
- [ ] Add Thunderbolt/USB-C port detection
- [ ] Export reports as PDF
- [ ] Multi-language support

### How to Contribute
1. Fork this repository
2. Create a feature branch (`git checkout -b feature/add-bios-check`)
3. Commit your changes (`git commit -m 'Add BIOS version detection'`)
4. Push to the branch (`git push origin feature/add-bios-check`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

## ⭐ Star This Repo

If this tool helped you, please give it a ⭐ — it helps others find it!

---

<p align="center">
  <b>Laptop Inspector</b> — Built with ❤️ for the IT community
</p>

