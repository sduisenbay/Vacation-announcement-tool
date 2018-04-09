# Vacation-announcement-tool
This is the Python code for one of my projects during 2nd rotation of DTLP program.   
The aim of the project is to create a tool to automate the manual **vacation announcement procedure**.   
The project automates the process of necessary Excel files generation and Outlook letter generation.  
    
The code uses _**pandas**_ library to work with Excel files, and _**win32**_ library to work with Outlook.
## Prerequisites for running the code && creating an executable file
#### Python installation
Installation of Python itself should be fairly straight-forward.
* Download and execute the latest Python 3.6.1 installation package from [here](https://www.python.org/downloads/release/python-361/)
* Put a tick when asked whether to add python path as environment variable
* Verify a successful installation by opening a command prompt window by typing `python` to open Python interpreter  

```Microsoft Windows [Version 10.0.15063]
(c) 2017 Microsoft Corporation. All rights reserved.

C:\Users\212633614>python
Python 3.6.1 (v3.6.1:69c0db5, Mar 21 2017, 18:41:36) [MSC v.1900 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>>
```
#### Pip installation
In order to install required libraries, you first have to install Pip.  
There are many methods for getting Pip installed, but my preferred method is the following:
* Download [get-pip.py](https://bootstrap.pypa.io/get-pip.py) to a folder on your computer. Open a command prompt window and navigate to the folder containing get-pip.py. Then run python get-pip.py. This will install pip.  
* Verify a successful installation by opening a command prompt window and typing 'pip freeze'  
pip freeze displays the version number of all modules installed in your Python non-standard library; On a fresh install, pip freeze probably won't have much info to show but we're more interested in any errors that might pop up here than the actual content



