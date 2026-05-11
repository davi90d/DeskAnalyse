Windows Hardware Diagnostic Tool
Description

Complete hardware diagnostic tool for Windows systems. Performs functionality tests on multiple computer components and generates detailed reports.

Features
✅ Intuitive graphical interface with organized tabs
✅ Automatic hardware detection using native Windows APIs
✅ Interactive testing for functionality validation
✅ HTML reports with professional formatting
✅ Standalone executable with no Python installation required
✅ Compatible with Windows 10/11
System Requirements
Operating System: Windows 10 or later
Privileges: Administrator rights (recommended for all tests)
Hardware: Webcam, microphone, and speakers (for specific tests)
Installation and Execution
Option 1: Precompiled Executable
Download the HardwareDiagnostic.exe file
Run it as administrator
Accept privilege requests when prompted
Option 2: Run with Python
# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
Option 3: Build Executable
# Install PyInstaller
pip install pyinstaller

# Run build script
python build_exe.py
Features and Functionalities
📊 Hardware Information
Motherboard: Manufacturer, model, serial number
Processor: Brand, model, specifications
RAM Memory: Total capacity, installed modules, speed
Storage Devices: Model, size, type
Graphics Card: Model, memory
Display: Resolution, display settings
TPM: Trusted Platform Module status
Connectivity: Bluetooth, Wi-Fi
🔧 Available Tests
1. Bluetooth Test
Detects Bluetooth adapters
Checks status and functionality
Lists paired devices
2. Keyboard Test
Interactive test for all keys
Real-time key press detection
Validation of special keys and combinations
3. TPM Test
Verifies Trusted Platform Module presence
Activation status
TPM version information
4. USB Test
Lists all connected USB devices
Detailed device information
Operational status check
5. Webcam Test
Real-time video capture
Resolution testing
Test image capture
6. Wi-Fi Test
Detects wireless network adapters
Connection status
Connected network information
7. Audio Test
Sound playback testing
Audio recording testing
Input and output device verification
📋 Reports
HTML Format: Professional reports with CSS styling
Text Export: Results in plain text format
Automatic Saving: Stored in the user's Documents folder
Timestamp: Date and time for every test
Project Structure
windows_version/
├── main.py                  # Application entry point
├── requirements.txt         # Python dependencies
├── build_exe.py             # Executable build script
├── README.md                # Documentation
├── core/                    # Core modules
│   ├── __init__.py
│   ├── hardware_info.py     # Hardware information collection
│   └── report_generator.py  # Report generation
├── gui/                     # Graphical interface
│   ├── __init__.py
│   └── main_window.py       # Main application window
└── tests/                   # Test modules
    ├── __init__.py
    ├── audio_test.py        # Audio testing
    ├── bluetooth_test.py    # Bluetooth testing
    ├── keyboard_test.py     # Keyboard testing
    ├── tpm_test.py          # TPM testing
    ├── usb_test.py          # USB testing
    ├── webcam_test.py       # Webcam testing
    └── wifi_test.py         # Wi-Fi testing
How to Use
1. Start the Application
Run HardwareDiagnostic.exe as administrator
The application will automatically verify dependencies
Accept privilege requests when prompted
2. View Hardware Information
Open the "Hardware" tab
Click "Refresh Information" if necessary
View detailed information about all components
3. Run Tests
Individual tests: Open the "Tests" tab → Click "Run" for each test
Run all tests: Click "Run All Tests"
User interaction: Some tests require user interaction
4. View Results
Open the "Results" tab
View detailed results for each test
Use "Save Results" to export
5. Generate Reports
Click "Generate Report"
HTML report will be created in the Documents folder
Automatically opens in the default browser
Dependencies
Required
tkinter - Graphical interface (included with Python)
Pillow - Image processing
opencv-python - Video capture
pyaudio - Audio processing
pynput - Keyboard event capture
numpy - Numerical computing
Optional (Recommended)
psutil - Detailed system information
py-cpuinfo - CPU information
WMI - Windows Management Instrumentation access
Development
pyinstaller - Executable generation
Troubleshooting
❌ Privilege Error

Problem: Some tests fail due to insufficient privileges
Solution: Run the application as administrator

❌ Webcam Not Detected

Problem: Webcam test fails
Solutions:

Verify the webcam is connected
Close other programs using the webcam
Update webcam drivers
❌ Audio Not Working

Problem: Audio test fails
Solutions:

Verify audio devices are working properly
Test audio with other applications
Check system volume
Reinstall audio drivers
❌ Incomplete Information

Problem: Some hardware information is missing
Solutions:

Run as administrator
Update system drivers
Some components may not be properly detected in virtual machines
❌ Missing Dependencies

Problem: Application fails due to missing libraries
Solution:

pip install -r requirements.txt
Limitations
Virtual Machines: Some components may not be properly detected
Legacy Hardware: Very old components may have limited detection support
Drivers: Information depends on installed drivers
Privileges: Some tests require administrative access
Development
Code Structure
Modular: Each test is an independent module
Object-Oriented: Dedicated classes for each functionality
Thread-safe: GUI remains responsive during tests
Error Handling: Friendly error capture and display
Adding New Tests
Create a new file inside tests/
Implement the class with methods initialize(), execute(), get_result()
Add it to the test list in main_window.py
Update the documentation
Build Process

The build_exe.py script automates the entire process:

Checks dependencies
Creates configuration files
Builds using PyInstaller
Generates installer and documentation
Version

1.0 - Windows Release

License

© 2024 - Hardware Diagnostic Tool
Davi Santos
Technical Analyst

Support

This tool is intended for basic diagnostics and hardware issue identification.
