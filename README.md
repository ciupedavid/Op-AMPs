# Step-by-step tutorial for installing Python, Selenium and Chromedriver using VS Code

Download and install VS Code from the official website: https://code.visualstudio.com/download
Download the latest version of Python from the official website https://www.python.org/downloads/
During the installation check the 'Add python.exe to PATH' then click Install Now
Open VS Code and go to the Extensions tab. Search for "Python" and install the Python extension by Microsoft. 
Run the installer and make sure to add Python to the system PATH during the installation process.
Open the Command Prompt and run the command `python --version` to check Python is installed correctly.
Check pip version by running `pip -V`, if it's not the latest version update pip using `python -m pip install --upgrade pip`
Use the command prompt and run `pip install selenium` to install Selenium.
Verify Selenium is installed correctly by running `pip show selenium` to check the version
Download the appropriate version of Chromedriver from the official website: https://chromedriver.chromium.org/downloads
Extract the downloaded file to a directory on your computer and note the location of the executable file.
Open the Command Prompt by searching for "Command Prompt" in the Start menu and clicking on the result.
Type `systempropertiesadvanced`  in the command prompt to open the environment variable editor
Click on the "Environment Variables" button.
Under "System variables," find the "Path" variable and click "Edit."
Click on "New" and add the path to the chromedriver executable. Click OK and close all the windows.
Create a new Python file in VS Code and start writing your Selenium code.
To run the Python code, use the built-in terminal in VS Code and run the command `python [filename.py]` or
    - press `F5` to run the script, or
    - click on the green run button on the top right corner, or 
    - use the command `Ctrl+Shift+B` to run the script
