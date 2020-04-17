#!/bin/bash
set -x

sudo yum install gcc-c++ pkgconfig poppler-cpp-devel python-devel redhat-rpm-config -y
sudo yum install poppler-utils -y
sudo yum install wget xz libjpeg-devel openjpeg2-devel -y
sudo yum install qt5-qtbase-devel -y
#sudo apt-get install python3-pyqt5
sudo yum install python-pip -y
python3 -m pip install --upgrade pip
python3 -m pip install pyqt5 --user
sudo -H pip install --upgrade -U pdfminer.six
sudo -H pip install --upgrade -U SETUPTOOLS
sudo -H pip install --upgrade -U pdftotext
sudo -H pip install --upgrade -U pdftotext
sudo -H pip install --upgrade -U pathlib
sudo -H pip install --upgrade -U pypdfocr
sudo yum update -y
sudo yum upgrade -y
sudo -H pip install --upgrade -U pycryptodome
sudo -H pip install --upgrade -U PyPDF2
sudo -H pip install --upgrade -U pdfrw
sudo -H pip install --upgrade -U Slate
sudo -H pip install --upgrade -U pytesseract
sudo -H pip install mkl-fft
