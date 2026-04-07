# MoveFilesFromExcel
Simple python script that reads the first column of an excel file containing file names and places said files into a custom directory

# Requirements
- Python 3.8+

# Clone Repository
First step to running this script it so clone this respository into your preffered IDE using this link https://github.com/NorthernSeasLTD/MoveFilesFromExcel

# Installation 
To install the required pacakages, run the following command pip install -r requirements.txt. Once the packages are download the script can be run with 3 specific arguments

1. First arg: path to the excel file that contains the file names in the first (or second) column
2. Second arg: path to the source directory to look for file names contained within the excel spreadsheet
3. Third arg: destination of the directory that the files need to be moved to 

## Command To Run
python ExtractFileToDir.py "excelDir" "sourceDir" "targetDir" 