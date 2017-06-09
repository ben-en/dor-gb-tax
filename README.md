# dor-gb-tax
A utility to convert Greenbits quarterly tax exports into Oregon Dept. of Revenue import compatible XLS files.

This rough guide assumes that you'll be running the script on a linux computer that experienced the same issues mine did (specifically, i was unable to open the Monthly Tax Report files downloaded from greenbits unless i opened the file in Microsoft Excel on windows and saved a copy into an XLS format.) No guarantees on windows/mac functionality.

Steps to acheive successful importing:
0a) Install the script on your linux computer:

    git clone https://github.com/ben-en/dor-gb-tax

    cd dor-gb-tax

    pip install -r requirements.txt

    pip install ./

 
0b) For each month, create a Monthly Tax Report from the first day to the last day of the month. As you download each file rename it to a day of the month, for example my report for March was renamed `3-1-17.xlsx`.

1) Put all files desired in the quarter in one folder.

2) On a windows computer open each file and save a copy to XLS format. <- Extremely critical step.

3) Send the files to the linux computer (i used email)

4) Run the script on the files all in the same folder.

    ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
    └╼ ls
    3-1-17.xls  4-1-17.xls  5-1-17.xls

    ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
    └╼ dor-gb-tax 3-1-17.xls; mv output.xls 1st-q.xls

    ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
    └╼ dor-gb-tax 4-1-17.xls 5-1-17.xls; mv output.xls 2st-q.xls


5) Send the output files back to the windows computer and save a copy to xlsx. This is equally important to the first time, as the DoR website requires xlsx.
