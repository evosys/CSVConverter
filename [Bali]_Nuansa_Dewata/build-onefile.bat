@REM @Author: ichadhr
@REM @Date:   2018-10-03 16:26:03
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-08 16:08:00

pyinstaller --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_Nusa_Dewata" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt app-nusa-dewata.py

pyinstaller --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_Nuansa_Dewata" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt app-nuansa-dewata.py

pyinstaller --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_Nuansa_Dewata_II" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt app-nuansa-dewata-II.py