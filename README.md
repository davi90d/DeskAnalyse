
# 🖥️ Windows Hardware Diagnostic Tool

A comprehensive hardware diagnostic tool for Windows systems. Performs functionality tests on various computer components and generates detailed reports.

![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?style=flat&logo=windows)
![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat&logo=python)
![License](https://img.shields.io/badge/License-Proprietary-red?style=flat)
![Version](https://img.shields.io/badge/Version-1.0-brightgreen?style=flat)

---

## ✨ Features

- **Intuitive graphical interface** with organized tabs
- **Automatic hardware detection** using native Windows APIs
- **Interactive tests** for component validation
- **Professional HTML reports** with clean formatting
- **Standalone executable** — no Python installation required
- **Compatible with Windows 10/11**

---

## ⚙️ System Requirements

| Requirement | Details |
|---|---|
| Operating System | Windows 10 or later |
| Privileges | Administrator (recommended for all tests) |
| Optional Hardware | Webcam, microphone, speakers (for specific tests) |

---

## 🚀 Installation & Usage

### Option 1: Pre-compiled Executable *(Recommended)*
1. Download `DiagnosticoHardware.exe`
2. Right-click → **Run as Administrator**
3. Accept privilege requests when prompted

### Option 2: Run via Python
```bash
# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```

### Option 3: Build Executable from Source
```bash
# Install PyInstaller
pip install pyinstaller

# Run build script
python build_exe.py
```

---

## 🔍 Functionality

### 📊 Hardware Information

Displays detailed information about the following components:

- **Motherboard** — Manufacturer, model, serial number
- **Processor** — Brand, model, specifications
- **RAM** — Total, installed modules, speed
- **Storage** — Model, size, type
- **GPU** — Model, video memory
- **Display** — Resolution, settings
- **TPM** — Trusted Platform Module status
- **Connectivity** — Bluetooth & Wi-Fi

---

### 🔧 Available Tests

| Test | Description |
|---|---|
| **Bluetooth** | Detects adapters, checks status, lists paired devices |
| **Keyboard** | Interactive real-time key detection, including special keys |
| **TPM** | Checks for Trusted Platform Module presence, activation and version |
| **USB** | Lists all connected USB devices with details and status |
| **Webcam** | Live video capture, resolution test, snapshot capture |
| **Wi-Fi** | Detects wireless adapters, connection status, network info |
| **Audio** | Tests playback and recording, verifies input/output devices |

---

### 📋 Reports

- **HTML format** — Professional reports with CSS styling
- **Text export** — Results in plain-text format
- **Auto-save** — Saved to the user's Documents folder
- **Timestamp** — Date and time for every test run

---

## 📁 Project Structure

```
windows_version/
├── main.py                 # Application entry point
├── requirements.txt        # Python dependencies
├── build_exe.py            # Executable build script
├── README.md               # This documentation
├── core/
│   ├── __init__.py
│   ├── hardware_info.py    # Hardware information collection
│   └── report_generator.py # Report generation
├── gui/
│   ├── __init__.py
│   └── main_window.py      # Main application window
└── tests/
    ├── __init__.py
    ├── audio_test.py
    ├── bluetooth_test.py
    ├── keyboard_test.py
    ├── tpm_test.py
    ├── usb_test.py
    ├── webcam_test.py
    └── wifi_test.py
```

---

## 📖 How to Use

**1. Launch the Application**
Run `DiagnosticoHardware.exe` as Administrator. The app will automatically verify dependencies.

**2. View Hardware Info**
Go to the **Hardware** tab and click **Refresh Information** to see detailed specs for all components.

**3. Run Tests**
- Run tests individually under the **Tests** tab by clicking **Run** next to each test.
- Or click **Run All Tests** to execute the full suite.
- Some tests require user interaction — follow on-screen prompts.

**4. View Results**
Open the **Results** tab to see detailed output for each test. Click **Save Results** to export.

**5. Generate Report**
Click **Generate Report** to create an HTML report. It will be saved to your Documents folder and automatically opened in your default browser.

---

## 📦 Dependencies

### Required
| Package | Purpose |
|---|---|
| `tkinter` | GUI framework (included with Python) |
| `Pillow` | Image processing |
| `opencv-python` | Video capture |
| `pyaudio` | Audio processing |
| `pynput` | Keyboard event capture |
| `numpy` | Numerical computation |

### Recommended
| Package | Purpose |
|---|---|
| `psutil` | Detailed system info |
| `py-cpuinfo` | CPU information |
| `WMI` | Windows Management Instrumentation access |

### Development
| Package | Purpose |
|---|---|
| `pyinstaller` | Build standalone executable |

---

## 🛠️ Troubleshooting

**Privilege Error**
Some tests fail due to insufficient permissions.
→ Run the application as Administrator.

**Webcam Not Detected**
→ Ensure the webcam is connected, close other apps using it, and update webcam drivers.

**Audio Not Working**
→ Verify audio devices are working in other applications, check system volume, and reinstall audio drivers if needed.

**Incomplete Hardware Info**
→ Run as Administrator. Some components may not be detected inside virtual machines.

**Missing Dependencies**
```bash
pip install -r requirements.txt
```

---

## ⚠️ Known Limitations

- **Virtual Machines** — Some components may not be detected correctly
- **Legacy Hardware** — Very old components may have limited detection
- **Drivers** — Displayed information depends on installed drivers
- **Privileges** — Some tests require Administrator access

---

## 🧱 Development

### Architecture
- **Modular** — Each test is an independent module
- **Object-Oriented** — Dedicated classes per feature
- **Thread-safe** — UI remains responsive during tests
- **Error Handling** — User-friendly error messages throughout

### Adding a New Test
1. Create a new file under `tests/`
2. Implement the class with `initialize()`, `execute()`, and `get_result()` methods
3. Register it in the test list inside `main_window.py`
4. Update the documentation

### Building the Executable
The `build_exe.py` script automates the full pipeline: dependency check → configuration → PyInstaller compilation → output generation.

---

## 👤 Author

**Davi Santos**  
Technical Analyst

---

## 📄 License

© 2024 — Hardware Diagnostic Tool. All rights reserved.
