SUMMARY
-----------------

This script counts the number of times a process is being used on a server and outputs it in to the results folder.

DEPENDANCIES
-------------------
It must run on a server with the Active Directory installed.

HOW TO
-------------

1. In the Config folder it gives 2 text files with an exclusion list. If you know some workers are offline short term, pop the name inside of the text file and it will exclude them from the check

2. Run the script by right clicking and "run with powershell".

3. It asks you for a process name. You don't need to specify the file type. So for "WinWord.exe" just put "WinWord".

4. Once finished it outputs the results to the Results folder in a file called ProcessCount.csv