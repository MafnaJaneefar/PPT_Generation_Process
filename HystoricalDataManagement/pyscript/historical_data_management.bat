@echo off
@REM echo %1
@REM echo %2
@REM echo %3
@REM echo %cd%
set current_directory=%cd%
set full_path=\pyscript
set python_env_path=\venv\Scripts\python.exe
set python_file_name=\app.py
@REM cd %current_directory%%full_path%
cd %current_directory%
%current_directory%%full_path%%python_env_path% %current_directory%%full_path%%python_file_name% "%1" "%2" "%3"
@REM pause
EXIT /B
