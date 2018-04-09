# Vacation-announcement-tool
This is the Python code for one of my projects during 2nd rotation of DTLP program.   
The aim of the project is to create a tool to automate the manual **vacation announcement procedure**.   
The project automates the process of necessary Excel files generation and Outlook letter generation.  
    
The code uses _**pandas**_ library to work with Excel files, and _**win32**_ library to work with Outlook.
## Prerequisites for running the code
### Python installation
Installation of Python itself should be fairly straight-forward.
* Download and execute the latest Python 3.6.1 installation package from [here](https://www.python.org/downloads/release/python-361/)
* Put a tick when asked whether to add python path as environment variable
* Verify a successful installation by opening a command prompt window and typing `python` to open Python interpreter  

```Microsoft Windows [Version 10.0.15063]
(c) 2017 Microsoft Corporation. All rights reserved.

C:\Users\212633614>python
Python 3.6.1 (v3.6.1:69c0db5, Mar 21 2017, 18:41:36) [MSC v.1900 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>>
```
### Pip installation
In order to install required libraries, you first have to install Pip.  
There are many methods for getting Pip installed, but my preferred method is the following:
* Download [get-pip.py](https://bootstrap.pypa.io/get-pip.py) to a folder on your computer. Open a command prompt window and navigate to the folder containing get-pip.py. Then run python get-pip.py. This will install pip.  
* Verify a successful installation by opening a command prompt window and typing `pip freeze`  
`pip freeze` *displays the version number of all modules installed in your Python non-standard library*  

```Microsoft Windows [Version 10.0.15063]
(c) 2017 Microsoft Corporation. All rights reserved.

C:\Users\212633614>pip freeze
chardet==3.0.4
click==6.7
Flask==0.12.2
future==0.16.0
```
### Libraries installation
Several packages should be installed before code can be executed.  
Open the command window and install the required libraries using `pip install ...`
1. Numpy - `pip install numpy`
2. Pandas - `pip install pandas`
3. PyInstaller - `pip install PyInstaller`
4. PypiWin32 - `pip install pypiwin32`
5. PyWin32 - `pip install pywin32`
6. Xlrd - `pip install xlrd`
7. Xlsxwriter - `pip install xlsxwriter`


## Creating an executable file
As the tool is to be used by employees with no programming background or knowledge, it is necessary to create an executable file. 
Currently, PyInstaller throws an error because of integration with _**pandas**_ library. The error is fixed by following this procedure:
1. Locate PyInstaller folder..\hooks, e.g. `C:\Users\212633614\AppData\Local\Programs\Python\Python36\Lib\site-packages\PyInstaller\hooks`
2. Create file `hook-pandas.py` with contents:  
`hiddenimports = ['pandas._libs.tslibs.timedeltas']`
3. Navigate to the folder where you downloaded `app.py`
4. Type `pyinstaller --onefile vacation_announcement_tool.py`  

The executable file will be created in the same directory in `dist` folder.


