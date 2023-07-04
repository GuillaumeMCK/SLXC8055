<div align="left">

```plaintext
-------------------------------------------------------------------------------
Date : 24/06/2021
Auteur : Guillaume MCK
-------------------------------------------------------------------------------
SYNOPSIS
Script to automate searching for a specific string in the log files of Xerox
C8055 printers.

DESCRIPTION
This script utilizes Selenium with ChromeDriver. The host machine must have
Google Chrome installed for the script to function properly.

To use the script, you need to create a file containing a list of printer names
or IP addresses to test (one printer per line) and specify the path to it as
an argument when running the script.

Example with a list: .\SGLX.ps1 "C:\Liste_C8055.txt"

PARAMETER InputPrinterName
The name of the printer or the path to the file containing a list of printer
names.

PARAMETER research
The string to search for in the log files.

PARAMETER inputList
Specify this switch if the InputPrinterName parameter represents a file
containing a list of printer names.

EXAMPLE
.\SGLX.ps1 "C:\Liste_C8055.txt" -research "apManager	crash"
```
</div>
