@echo off
F:

timeout /t 5 /nobreak
cd PSTs

timeout /t 5 /nobreak
set /p PstFolder=Enter the path to the PST folder:

timeout /t 5 /nobreak
set /p URL_SAS=Enter the SAS URL:

timeout /t 5 /nobreak
azcopy.exe copy "%PstFolder%" "%URL_SAS%" --recursive
