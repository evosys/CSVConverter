pyinstaller --debug --onefile --clean --noconsole --noupx --icon=resources\\icon.ico --name="CSV_Converter" --key=68b00c755cef892e512d56621925d836 --version-file=version.txt --add-binary "tabula/tabula-1.0.2-jar-with-dependencies.jar;tabula/" app.py