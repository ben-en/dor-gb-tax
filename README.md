# dor-gb-tax
A utility to convert Greenbits quarterly tax exports into Oregon Dept. of Revenue import compatible XLS files.

This rough guide assumes that you'll be running the script on a linux computer that experienced the same issues mine did (specifically, i was unable to open the Monthly Tax Report files downloaded from greenbits unless i opened the file in Microsoft Excel on windows and saved a copy into an XLS format.) No guarantees on windows/mac functionality.

  Steps to acheive successful importing:
  
  1) Install the script on your linux computer:
  
```
      git clone https://github.com/ben-en/dor-gb-tax

      cd dor-gb-tax

      pip install -r requirements.txt

      pip install ./
```
 
  2) For each month, create a Monthly Tax Report from the first day to the last day of the month. As you download each file rename it to a day of the month, for example my report for March was renamed `3-1-17.xlsx`.

  3) Put all files desired into one folder.

  4) On a windows computer open each file and save a copy to XLS format. <- Extremely critical step.

  5) Send the files to the linux computer (i used email)

  6) Run the script on the files all in the same folder. This will generate
  output for all quarters you've provided months for.

```
      ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
      └╼ ls
      3-1-17.xls  4-1-17.xls  5-1-17.xls

      ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
      └╼ dor-gb-tax 3-1-17.xls 4-1-17.xls 5-1-17.xls

      ┌──(0) ben [~/code/tax-conversion/dor_gb_tax/example]
      └╼ ls
      1st-Quarter.xls  2nd-Quarter.xls  3-1-17.xls       4-1-17.xls       5-1-17.xls
```

  7) Send the output files back to the windows computer and save a copy to xlsx. This is equally important to the first time, as the DoR website requires xlsx.
