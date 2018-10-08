@REM @Author: ichadhr
@REM @Date:   2018-10-02 17:28:58
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-08 14:26:20

pyinstaller --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt --add-data="tabula/bin/tabula-1.0.2-jar-with-dependencies.jar;tabula/bin/" app.py