# checkWhoModifiedMy_Xlsx
This is a simple PowerShell script that can be run locally in order to check who last modified a given .xlsx file.

I am proud of this project because it showcases my problem solving capabilities. This PowerShell script was not easy to figure out in any sense of the word - I had to come up with new, and perhaps odd, methods to achieve my goal. Microsoft never makes it easy to do anything within any of their products; at least, that's what I've learned. 
In order to find who last modified a .xlsx file, I had to figure out that .xlsx files contain core.xml files, which hold properties such as modifiedBy. You can only access this core.xml file if you unzip the target file first, so I used the System.IO.Compression namespace to perform this operation.
So - I used that function to unzip the file into a new temp location, which I would then delete after I'm finished extracting the properties I need from the .xml. 
