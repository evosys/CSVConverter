@REM @Author: ichadhr
@REM @Date:   2018-06-06 11:17:38
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-03 09:52:09

pyinstaller --debug --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt --add-binary "tabula/tabula-1.0.2-jar-with-dependencies.jar;tabula/" app.py