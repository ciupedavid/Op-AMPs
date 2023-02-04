# Step-by-step tutorial for installing Python, Selenium and Chromedriver using VS Code

1.	Download and install VS Code from the official website:
a.	https://code.visualstudio.com/download
2.	Download the latest version of Python from the official website
a.	https://www.python.org/downloads/
b.	Add Python to the PATH during the installation process.
3.	Download MS Visual C++ 14.0 or greater:
a.	Go to https://visualstudio.microsoft.com/visual-cpp-build-tools/
b.	Click on 'Download Build Tools'
c.	Run 'vs_BuildTools.exe'
d.	Select Dextop developement with C++ from Dextop & Mobile section
e.	Select only the MSVC v143 - VS2022 C++ x64/x86 and Windows 11 SDK
4.	Install Python extention in Visual Studio Code
a.	Go to the Extentions from the left side menu (Ctrl+Shift+X)
b.	Search for "Python" and install the Python extension.
5.	Select Python Interpreter:
a.	Go to Command Palette or from the View tab (Ctrl+Shift+P) 
b.	Type Python: Select Interpreter then select Python 3.x from the dropdown
6.	Verify is Python and Pip are installed
a.	Open Terminal and run `python --version` to check Python is installed correctly.
b.	Check pip version by running `pip -V`, if it's not the latest version update pip using `python -m pip install --upgrade pip`
7.	Installing python libraries/requirements
a.	Download the requirements.txt in the Python project folder
b.	Run in terminal: pip install -r requirements.txt
8.	(Optional) Download the appropriate version of Chromedriver from the official website:
a.	Check Chrome Browser version from Settings > Help > About Google Chrome (v109.x.x.x)
b.	Go to: https://chromedriver.chromium.org/downloads
c.	Download the version matching the Chrome Browser version (v109)
d.	Extract the downloaded file to a directory on your computer and note the location of the executable file.
e.	Open the Command Prompt by searching for "Command Prompt" in the Start menu and clicking on the result.
f.	Type `systempropertiesadvanced`  in the command prompt to open the environment variable editor
g.	Click on the "Environment Variables" button.
h.	Under "System variables," find the "Path" variable and click "Edit."
i.	Click on "New" and add the path to the chromedriver executable. Click OK and close all the windows.
9.	Create a new Python file in VS Code and start writing your Selenium code.
10.	To run the Python code, use the built-in terminal in VS Code and run the command `python [filename.py]` or
    - press `F5` to run the script, or
    - click on the green run button on the top right corner, or 
    - use the command `Ctrl+Shift+B` to run the script
1.	Open the Command Prompt and run the command `python --version` to check Python is installed correctly.
2.	Check pip version by running `pip -V`, if it's not the latest version update pip using `python -m pip install --upgrade pip`
3.	Verify Selenium is installed correctly by running `pip show selenium` to check the version
4.	

