SUMMARY
-----------------

GET-WorkerEventLogs gets the newest 60 errors in the event logs of all worker servers.

DEPENDANCIES
-------------------
It must run on a server with the Active Directory installed.

HOW TO
-------------

1. In the Config folder it gives 2 text files with an exclusion list. If you know some workers are offline short term, pop the name inside of the text file and it will exclude them from the check

2. Run the script by right clicking and "run with powershell".

3. Once finished it outputs the results to the Results folder in files called Farm1_error_logs.csv and Farm2_error_logs.csv