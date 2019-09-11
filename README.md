# wcexport - WebChart Export Utility

This standalone app allows webchart users to use the Print Chart feature of WebChart to export pdf charts from the system.

Requirements
----------------
* A url to the WebChart system. This url should include the 'webchart.cgi' part.
* A print definition to use for printing all charts
* A system report that queries the system for charts to be exported
* Valid WebChart credentials
* 'Appliance Synchronization' permission granted

System Report
----------------
The system report should contain two columns:
* pat_id - This should be the patient/chart internal pat_id identifier. Each pat_id in WebChart corresponds to a single unique patient/chart so if your report has multiple joins, ensure that you group by pat_id
* filename - This should the be the desired resulting filename of the exported chart. Some escaping may be done to ensure that the filename is valid, but you should take note to try to avoid weird characters and be wary of your local filesystem's maximum filename length limits.

Running
--------------
To run the executable on Windows, simply run the `wcexport.exe` file.

You may also run the python script directly if you cannot run a Windows executable. `./wcexport.py`


Enter the required information into the application's fields and click 'Begin Export'.

The printed charts will be downloaded to the user's home directory in the wcexport folder.

