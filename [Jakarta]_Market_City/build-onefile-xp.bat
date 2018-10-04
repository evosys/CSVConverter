@REM @Author: ichadhr
@REM @Date:   2018-10-02 12:02:05
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-02 12:02:44

pyinstaller --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter" --version-file=version.txt app.py