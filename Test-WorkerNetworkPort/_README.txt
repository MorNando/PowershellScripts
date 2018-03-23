SUMMARY
-----------------

Test-WorkerNetworkPort tests the connection from all worker servers to an ip address of your choosing using a specific port number.

DEPENDANCIES
-------------------
It must run on a server with the Active Directory installed.

HOW TO
-------------

1. In the Config folder it gives 2 text files with an exclusion list. If you know some workers are offline short term, pop the name inside of the text file and it will exclude them from the check

2. Run the script by right clicking and "run with powershell".

3. Once finished it outputs the results to the Results folder in a file called FaultyPortServers.txt