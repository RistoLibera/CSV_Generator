# Project: CSV_Generator
## What does this achieve

- Generate csv file by Excel file data

## Getting start

- Put vbs(.vbs) script and xlsx file into the same folder
- Open command prompt on the same folder
- In command prompt run command below to generate an csv file(thisisnotcsv.csv) from range[A1:I10] in "data" sheet 
```cmd
.\CSV_Generator.vbs A1 I10 data thisisnotcsv
```

## Notes

- The script will detect the existence of Excel file. If there are more than one Excel file, an error may pop up