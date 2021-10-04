# ListPOR

The ListPOR program (List Parser for Outlier Removal) can be used to read a file containing columns of grouped values 
and remove outlier values using Grubb's test. 

## Usage

Typically, the first column of the input file will contain key information (text or numbers) 
on which to group the data. The second column will contain the values to examine for outliers, 
examining the data by groups. If additional columns are present, they will be written to the 
output file along with the key and value columns. Alternatively, if you provide a file with only 
one column of data, then all of the values will be examined en masse and the outliers removed.

If you check "Use Symmetric Values", the data will be converted to symmetric
values before processing.  This option is appropriate for abundance ratio data that
is centered around 1 (i.e. 1 is unchanged, >1 is up-regulated, <1 is down-regulated,
all data is > 0).  The logic used to convert to convert each value in the input file
to a symmetric value is:

* If the value is >= 1, compute: result = value - 1
* If the value is > 0 and < 1, compute: result = 1 - 1 / value
* If the value is <= 0, compute: result = 0

The symmetric values will then be used for looking for outliers.  The output file
will contain the original, unaltered values.

If you check "Use Natural Log Values", then the data will be converted to natural
log equivalents before processing.  This is an alternative method for dealing with
abundance ratio data.

If the input file contains a key column and value columns, then, after outlier 
removal the values remaining for each key will be averaged together and the average 
written as an additional column in the output file.  If Symmetric Values were used,
then the average will be computed using the symmetric values, and the result
converted back to a normal value before being written.

## Installation

* Download ListPOR_Setup.zip from [GitHub](https://github.com/PNNL-Comp-Mass-Spec/ListPOR/releases)
* Extract the files
* Run ListPOR_setup.exe

## Console Switches

ListPOR can be run in batch mode from the Windows command line.  Syntax:

```
ListPOR.exe 
  /I:InputFilePath [/O:OutputFilePath] [/Sorted] [/L] 
  [/M:MinimumFinalDataPointCount] [/C:ColumnCount] [/Conf:#]
```

The input file is a tab-delimited data file with one or more columns of data; specify it using 
/I or by simply using the file path (use double quotes if spaces in the path)

The output file path is optional. If omitted, the output file will be created in the same folder 
as the input file, but with _filtered appended to the name

/Sorted indicates that the input file is already sorted by group. This allows very large files 
to be processed, since the entire file does not need to be cached in memory. It is also useful 
for obtaining a filtered file with data in the exact same order as the input file.

/L will cause the program to convert the data to symmetric values, prior to looking for outliers. 
This is appropriate for data where a value of 1 is unchanged, >1 is an increase, and <1 is a decrease. This is not appropriate for data with values of 0 or less than 0.

/M is the minimum number of data points that must remain in the group. It cannot be less than 3

/C:1 can be used to indicate that there is only 1 column of data to analyze; if other columns 
of text are present after the first column, they will be written to the output file, 
but will not be considered for outlier removal. Use /C:2 to specify that there are two columns 
to be examined: a Key column and a Value column. Again, additional columns will be written to disk, 
but not utilized for comparison purposes. 

/Conf can be used to specify the confidence level; options are
* /Conf:95
* /Conf:97
* /Conf:99

## Contacts

Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) \
E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov\
Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics\
Source code: https://github.com/PNNL-Comp-Mass-Spec/ListPOR

## License

ListPOR is licensed under the Apache License, Version 2.0; you may not use this 
file except in compliance with the License.  You may obtain a copy of the 
License at https://opensource.org/licenses/Apache-2.0
