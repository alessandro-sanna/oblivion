# Oblivion
Oblivion is a python3-based framework for automatically analysing Office files containing VBA macros. It can only be run on Windows. See "Installation" for further detail.

## Usage

Oblivion is designed for the analysis of multiple files. To analyze multiple files at once, run the following:

`python3 ./oblivion.py -d /path/to/office/files -o /path/where/to/save/reports`

You can also analyze one single file:

`python3 ./oblivion.py -f /path/to/office/file -o /path/where/to/save/report`
```
usage: oblivion.py [-h] (-f [TARGET_FILE] | -d [TARGET_DIRECTORY]) -o
                   [OUTPUT_DIRECTORY] [-t [TIME_LIMIT]] [-mdb] [-ncs] [-dd]
                   [-nt [MAX_RETRIES]] [-sf [START_FROM]]

options:
  -h, --help            show this help message and exit
  -f [TARGET_FILE], --file [TARGET_FILE]
                        path of a single to-be-analyzed file, cannot be used
                        with -d
  -d [TARGET_DIRECTORY], --directory [TARGET_DIRECTORY]
                        path of a directory of to-be-analyzed files, cannot be
                        used with -f
  -o [OUTPUT_DIRECTORY], --output [OUTPUT_DIRECTORY]
                        path of the directory where Oblivion will save the
                        report file
  -t [TIME_LIMIT], --time_limit [TIME_LIMIT]
                        maximum time per single analysis
  -mdb, --use_mongo_db  if set, save reports in a mongo database (NOT IMPLEMENTED YET!)
  -ncs, --no_clean_slate
                        if set, inject instrumentation in file as it is
  -dd, --in_depth       if set, look recursively in subdirectories
  -nt [MAX_RETRIES], --max_retries [MAX_RETRIES]
                        if set, the file can try to run again NT times after a VBA
                        exception (NOT IMPLEMENTED YET!)
  -sf [START_FROM], --start_from [START_FROM]
                        skip first N samples in folder(s)
```


## Installation
First, install requirements:

`pip install -r requirements.txt`

Then, add keys to Windows Registry:
- DWORD \HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Excel\Security\blockcontentexecutionfrominternet 0
- DWORD \HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Excel\Security\vbawarnings 0
- DWORD \HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Word\Security\blockcontentexecutionfrominternet 0
- DWORD \HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Word\Security\vbawarnings 0

This program requires an installation of Microsoft Office (https://www.office.com/) and Sandboxie (https://sandboxie-plus.com/). Follow the respective vendors' instructions. We also strongly advise running this framework inside a Virtual Machine if you are dealing with malware.

In Word and Excel: go to Trust Center and allow Programmatic Access to the VBA Module; tick "Enable all Macros".

In Sandboxie: disable all alerts.

Finally, set the paths required in $OBLIVIONFOLDER/OblivionResources/config/configuration.json

You should be good to go!
