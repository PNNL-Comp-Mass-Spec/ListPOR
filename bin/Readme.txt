ListPOR (List Parser for Outlier Removal)

The ListPOR program can be used to read a file containing columns of grouped values 
and remove outlier values using Grubb's test. Typically, the first column of the
input file will contain key information (text or numbers) on which to group the data. 
The second column will contain the values to examine for outliers, examining the 
data by groups. If additional columns are present, they will be written to the 
output file along with the key and value columns. Alternatively, if you provide a 
file with only one column of data, then all of the values will be examined en masse 
and the outliers removed.

If you check "Use Symmetric Values", then the data will be converted to symmetric
values before processing.  This option is appropriate for abundance ratio data that
is centered around 1 (i.e. 1 is unchanged, >1 is up-regulated, <1 is down-regulated,
all data is > 0).  The logic used to convert to convert each value in the input file
to a symmetric value is:

If the value is >= 1, compute: result = value - 1
If the value is > 0 and < 1, compute: result = 1 - 1 / value
If the value is <= 0, compute: result = 0

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

Double click the ListPOR_Installer.msi file to install.  The program
shortcut can be found at Start Menu -> Programs -> PAST Toolkit -> ListPOR

-------------------------------------------------------------------------------
Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
Program started in August 2004

E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

All publications that result from the use of this software should include 
the following acknowledgment statement:
 Portions of this research were supported by the W.R. Wiley Environmental 
 Molecular Science Laboratory, a national scientific user facility sponsored 
 by the U.S. Department of Energy's Office of Biological and Environmental 
 Research and located at PNNL.  PNNL is operated by Battelle Memorial Institute 
 for the U.S. Department of Energy under contract DE-AC05-76RL0 1830.

Notice: This computer software was prepared by Battelle Memorial Institute, 
hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
Department of Energy (DOE).  All rights in the computer software are reserved 
by DOE on behalf of the United States Government and the Contractor as 
provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
SOFTWARE.  This notice including this sentence must appear on any copies of 
this computer software.
