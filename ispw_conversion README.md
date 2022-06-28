# Convert ISPW deployment configuration contents

Typically a deployment requires a number of different resouces to be deployed depending on the environment, sub-environment, domain, type, etc. When viewing the definitions in ISPW it is not straight forward to compare against a master list, etc becuase a set of definitions is formated vertically over myltiple rows.
This Excel tool is written to format the exported contents from ISPW into one line by extracting the following four fields.

* DPENV
* SUBENV
* DPTYPE
* DPDOMAIN
* STORNAME

How to use
Split the exported contents with "=" as a delimiter in Excel.
Copy the entire data to "Sheet4" or in the sheet defined in the VBA source code.
Run the macro.
Check "Sheet2" or whatever the sheet defined in the VBA source code.
 
Note: Becuase the current logic relies on STORNAME to complete a new line there may be an incomplete converted line at the end. This occurs if the last set of configuration does not have STORNAME. Please ignore or manually delete the last line if required. 
