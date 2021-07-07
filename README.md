# Data Cleanser for QSDR

<img alt="DataCleanserGUI" src="https://i.imgur.com/WIYpykZ.gif"></img>

## Purpose of Application 
This is a Python application for removing errors in a specific type of excel documents.

Specifically, it is used to remove errors in _Quarterly Skills Development Reports_ (QSDR).  These are records of skills programmes financed by the organization for a financial quarter. 

A variant of this application was developed for a large organization. These reports contain a great amount of data which were manually updated and verified. Consequently, gross errors were frequent resulting in adverse financial implications.  

This application was developed to:
- Automatically fix data errors using baseline data
- Provide quantitative data on errors 
 
### Business benefit
Data Cleanser negated the need to manually validate data, which was conducted by teams of people. Not only did this process take multiple hours to complete, but it also occasionally exacerbated issues. 
 
 
## Getting started
To run this application, in addition to installing [Python](https://www.python.org/), the following libraries are required: 
- [Matplotlib](https://matplotlib.org/)
- [Pandas](https://pandas.pydata.org/)
- [Numpy](https://numpy.org/)
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [xlrd](https://pypi.org/project/xlrd/)

_Data Cleanser_ was designed to be a portable desktop application. You may convert it an executable using the [Pyinstaller]( https://www.pyinstaller.org/) command:
````
pyinstaller --onefile -w --icon=Image_files\Icon0.ico main.py --version-file file_version.txt
````
**Note**: The icon and version file are included in the repository. 


## How to use
A detailed guide on how to use this application can be found [here](Documents/User_Guide.pdf)


## Contact
Zaahier Adams – ZaahierAdams2021@gmail.com

Project link: https://github.com/ZaahierAdams/Data-Cleanser-for-QSDR 


