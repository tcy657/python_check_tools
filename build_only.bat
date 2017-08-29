rd /s /q build
del /s /q *.spec
del /s /q *.exe


rem copy "C:\Users\fh\Desktop\exePathTest.py"  .\exePathTest.py
python %pyinstaller%  -F  -c --icon="otnm_clear.ico" startCheck.py --distpath ./
pause