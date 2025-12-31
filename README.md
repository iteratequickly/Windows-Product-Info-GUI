# 🛡️ Windows Product Information GUI (PowerShell + WPF)

A modern **dark-themed Windows GUI tool** built with **PowerShell and WPF** that displays detailed Windows product and system information: including decoded product keys and installation IDs: with one-click copy functionality.

> **Disclaimer:** This tool does not activate Windows or modify licensing. It only displays locally available system information using official Windows interfaces.

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## ✨ Features

- 🔍 **Detailed OS Info:** Displays OS Name, Edition, OS Build (including UBR), and Architecture.
- 🔑 **Advanced Key Decoding:** Uses a character map (BCDFGHJKMPQRTVWXY2346789) and bit-wise operations to decode the binary `DigitalProductId` from the registry.
- 📋 **Clipboard Support:** Dedicated "Copy" buttons for every field, including a formatted version of the Installation ID.
- 🧠 **Multi-Method Retrieval:** Combines Registry decoding with CIM/WMI fallbacks to ensure information is captured even on different license types.
- 🖥️ **Modern UI:** Responsive dark-themed WPF GUI with high-contrast text and "toast" notifications when data is copied.
- ⚡ **Performance Optimized:** Features CIM session caching, batched registry reads, and pre-compiled regex patterns.
- 🛡️ **Self-Elevation:** Automatically detects if it is running without privileges and prompts to restart as Administrator.

---

## 🖼️ Displayed Information

- **System:** OS Name, Edition ID, OS Build Version (e.g., 19045.5247), and local Installation Date.
- **Licensing:** Activation Status, Decoded Product Key, and License Channel (Retail, OEM, or Volume/KMS).
- **Activation Support:** Retrieves the 63-digit **Installation ID** (IID) via background jobs to keep the UI responsive, displaying it in 9 distinct groups for phone activation.

---

## 🚀 Getting Started

### Prerequisites
- Windows 10 or Windows 11.
- PowerShell 5.1 or newer.
- Administrator privileges (the script will prompt automatically).

### Running the Tool
1. **Download** `WindowsProductInfo.ps1`.
2. **Right-click** the script file in your folder.
3. Select **Run with PowerShell**.
4. If prompted by User Account Control, click **Yes** to allow Administrator elevation.

---

## 🛠️ Technical Background
This tool implements logic based on community-developed methods for decoding the Windows DigitalProductId. Special thanks to the researchers and developers at:
- PowerShell.one
- Learn-PowerShell.net
- mrpear.net (WinProdKeyFinder)
- chentiangemalc.wordpress.com

## ⚖️ License
This project is licensed under the **MIT License**, see the [LICENSE](LICENSE) file for details.
