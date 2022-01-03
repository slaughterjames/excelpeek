## ExcelPeek v0.2 - Installation Guide

ExcelPeek has been tested on Ubuntu-based Linux distributions.  It is generally recommended using the the latest available version (or at least the latest long-term supported version).

## Download the ExcelPeek Files

1. Download the ExcelPeek files:
```bash
wget https://github.com/slaughterjames/excelpeek/archive/refs/heads/main.zip
```
OR

```bash
gh repo clone slaughterjames/excelpeek
```

2. Grant execution privileges:
```
chmod +755 excelpeek.py
```

## Install Python Prerequisites

1. Using PIP:

```bash
sudo pip3 install re
```
```bash
sudo pip3 install termcolor
```
```bash
sudo pip3 install openpyxl
```

