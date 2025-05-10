# PLC Data Writer V3

## 📋 Description

**PLC Data Writer V3** is a Python application designed to interface with Rockwell Automation PLCs and write specific data to Excel files using a graphical interface.
It is intended for **authorized personnel** and designed for **use with a specific machine and table format**. Unauthorized use may result in **equipment malfunction or personnel injury**.

> ⚠️ **Warning:** This tool directly communicates with industrial equipment. Use with caution.

---

## 🖥️ Features

- Real-time connection to Rockwell PLCs using Ethernet/IP.
- Read data from structured Excel templates.
- Simple, intuitive GUI using Tkinter.
- Optional image/logo integration.

---

## 🚀 Getting Started

### Prerequisites

- Python 3.12
- Installed dependencies (see below)

### Installation

```bash
pip install pycomm3 openpyxl Pillow
```
Tkinter is included in most standard Python installations. If you're missing it on Linux:

```bash
sudo apt-get install python3-tk
```

---

▶️ Usage
Run the application using:

```bash
python main.py
```
Fill in required fields.

Click "Write to PLC"

---

🙏 Special Thanks
Special thanks to the authors and maintainers of the following libraries used in this project:

pycomm3 – For Rockwell PLC communication in Python.

openpyxl – For reading and writing Excel files.

Pillow – For image handling.

Tkinter – For GUI components (part of the Python standard library).
