@echo off

:: Change directory to the project folder
cd D:\your-project-name\familyfund\rjsiao_family_fund


:: Get the current date and extract the current year and month
for /f "tokens=2 delims==" %%I in ('"wmic OS Get localdatetime /value"') do set datetime=%%I
set curr_year=%datetime:~0,4%
set curr_month=%datetime:~4,2%
set curr_yyyymm=%curr_year%%curr_month%

echo "Current yyyy-mm today: %curr_yyyymm%"

:: Prompt the user to enter the first argument (yyyymm)
set /p month1="Enter start yyyymm (e.g. 202501): "

:: Prompt the user to enter the second argument (yyyymm)
set /p month2="Enter end yyyymm (e.g. 202503): "

:: Swap month1 and month2 if month1 is larger than month2
if %month1% gtr %month2% (
    set mon = %month1%
    set month1 = %month2%
    set month2 = %mon%
    echo month1 was larger than month2. Swapping values...
)

:: Display the final range after swapping (using delayed expansion)
echo "yyyy-mm Range for download: from %month1% to %month2%"

:: Call the Python script with the two user inputs as arguments
python merge_fund_balance.py %month1% %month2%

:: Pause the execution so the user can see the results
pause