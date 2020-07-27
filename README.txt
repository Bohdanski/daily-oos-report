# Daily OOS Report

Daily OOS Report is a package that consilidates IDW OOS reports into the main out of stock report.



# Installation

This program is mostly self contained within its folder. Where-ever you decide to place the folder, the path must be manually adjusted within the launcher.bat script.



# Usage

step 1 - Save all files to .\\excel\\data\\
step 2 - Make sure the path is correct within the "launcher.bat" file
step 3 - Double click the "launcher.bat" file



# Versions

v0.1 - 05/01/2020 - Beta release that compiles all files but is currently unable to append the files to the template without corrupting the Excel file.
v0.2 - 07/27/2020 - Script now handles the file from CS.
		  - Improved error handling format, accounts for a missing CS file.
		  - Script will delete Excel files within the data directory at the end of the proccess.
		  - Added simple feedback to the terminal to see what the script is doing at various stages for debugging purposes.



# Author

Bohdan Tkachenko