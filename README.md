**WHAT IT DOES.**

Python 3 script which look through all files and folders starting with the root directory which you have assigned.<br />
Parses: .doc .docx .xls .xlsx .pdf .msg .txt<br />
Outputs: .txt file with a list of all found files, aswell as the line number of the regex match and the match itself.

**HOW TO.**

0. Install the required libraries, I highly recommend to do this on linux.
1. Assign the root directory to the **path** variable
2. Assign the output file location to the **output_path** variable
3. Assign the output file location for errors to the **error_path** variable
4. Run using python3

**LATEST UPDATE : February 21st 2019**

Redesigned the core functions for PDF and XLS parsing. The main regex has also been redone, just all around changes in pretty much every function.Results include no more false positives and a more fleshed out output file.
