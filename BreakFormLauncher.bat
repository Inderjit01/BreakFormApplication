@echo off
:: Disable command echoing to make the output cleaner

:: Check if Python is installed
:: Attempt to get the python version; redirect output to null to suppress it
python --version >nul 2>&1

:: Check the error level to see if the previous command failed
if %errorlevel% neq 0 (
    echo Python is not installed. Downloading and installing Python...

    :: Download Python installer (change the URL to the latest version if needed)
    :: Use curl to download the installer
    curl -o python-installer.exe https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe
	
    :: Check if the download was successful
    if exist python-installer.exe (
        :: Install Python (adjust arguments if needed)
        start python-installer.exe InstallAllUsers=1 PrependPath=1
		@echo installing python please wait. Click anywhere when done installing
		pause

        :: Clean up the installer file after installation
        del python-installer.exe
		@echo Deleted installer
		@echo THE PROGRAMS NEEDS TO RESTART PLEASE RELAUNCH AFTER CONTINUING
		:: Open the Rickroll after installing in the default web browser
		start https://www.youtube.com/watch?v=dQw4w9WgXcQ
		pause
		exit
		
    )
) else (
    echo Python is already installed.
)

:: Check if python and pip are installed
Python --version
pip --version

:: check if addiional libraries are installed if not install them

pip show PyQt5 >nul 2>&1
if %errorlevel% neq 0 (
    echo PyQt5 not found. Installing...
    pip install PyQt5
)

pip show docx2pdf >nul 2>&1
if %errorlevel% neq 0 (
    echo docx2pdf not found. Installing...
    pip install docx2pdf
)

pip show python-docx >nul 2>&1
if %errorlevel% neq 0 (
    echo python-docx not found. Installing...
    pip install python-docx
)

pip show beautifulsoup4 >nul 2>&1
if %errorlevel% neq 0 (
    echo beautifulsoup4 not found. Installing...
    pip install beautifulsoup4
)

@echo Searching for program...
:: Define the folder and Python script names
set "folder_name=BreakFormApplication"
set "script_name=main.py"

:: Search the C: drive for the folder and script
for /d /r "C:\Users" %%d in (*) do (
    if /i "%%~nxd"=="%folder_name%" (
        if exist "%%d\%script_name%" (
            echo Found script at: %%d\%script_name%
            cd /d "%%d"
            python "%script_name%"
            goto :script_found
        )
    )
)

echo Folder "%folder_name%" with the script "%script_name%" not found in C drive.
echo Searching D drive...
:: Search the D: drive for the folder and script
for /d /r "D:\" %%d in (*) do (
    if /i "%%~nxd"=="%folder_name%" (
        if exist "%%d\%script_name%" (
            echo Found script at: %%d\%script_name%
            cd /d "%%d"
            python "%script_name%"
            goto :script_found
        )
    )
)
:script_found
pause
