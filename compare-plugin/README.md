# CompareResults

The configuration assessment tool (CAT) see [here](https://github.com/Appdynamics/config-assessment-tool) provides metrics on how well your applications are instrumented based on some of the field best practices.

The CompareResults project piggy backs on to the output of the CAT, by allowing us to compare previous output against current output for APM only - these are workbooks ending with "-MaturityAssessment-apm.xlsx"

## Requirements

- Python 3.x

# Setup!

## Setup Instructions

1. Unzip the `CONFIG-ASSESSMENT-TOOL` folder.

2. On mac - open a terminal and navigate to the `compare-plugin` directory:
    cd path/to/unzipped/CONFIG-ASSESSMENT-TOOL/compare-plugin

3. Run the setup script using bash:
    ./setup.sh

4. After the bash script has complete and if all modules have been installed - run the following commands 1 after the other:
    source venv/bin/activate
    python3 core.py

5. The UI should automatically launch with an address of: http://127.0.0.1:5000/ - see considerations for upload. 
    - The only CAT report we can compare at this time is the APM output - ending with "-MaturityAssessment-apm.xlsx" 
    - The previous and current APM report has to be from the same controller - otherwise the script will terminate
    - For best results ensure the previous APM report is dated before the Current APM report



## If bash will not run:

- chmod +x setup.sh

## Module Not Found Errors
Modules should be installed as part of setup.sh, however, if you get Module Not Found Errors when running core.py (Error: ModuleNotFoundError: No module named 'openpyxl') you will have to install each Module.

Below is a list of the modules needed:
- Flask
- pandas
- openpyxl
- python-pptx
- xlwings

Below is the versions being used
- Flask>=2.3.2
- pandas>=1.5.3
- openpyxl>=3.1.2
- python-pptx>=0.6.21
- xlwings>=0.28.0

Install Module as follows:
- pip3 install <<module>> or pip install <<module>>
- Example if you receive: Error: ModuleNotFoundError: No module named 'openpyxl' - enter pip3 install openpyxl 


To help with a successful comparison, see below: 
1. Only one workbook "{jobName}-MaturityAssessment-apm.xlsx" can be compared for now
2. The xlsx files to be compared have to be from the same controller 