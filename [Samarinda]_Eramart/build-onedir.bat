@REM @Author: ichadhr
@REM @Date:   2018-07-11 10:50:15
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-08 16:53:57

pyinstaller --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_era-500-lembuswana" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt era_500_lembuswana.py


pyinstaller --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_eramart-pelita" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt eramart_pelita.py


pyinstaller --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_eramart-tengkawang" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt eramart_tengkawang.py


pyinstaller --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter_eramart-suryanata-1" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt eramart_suryanata_1.py