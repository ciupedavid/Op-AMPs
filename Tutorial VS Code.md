# Step-by-step tutorial for installing Python, Selenium and Chromedriver using VS Code

1. Download and install VS Code from the official website: https://code.visualstudio.com/download
2. Open VS Code and go to the Extensions tab. Search for "Python" and install the Python extension by Microsoft. 
3. Download the latest version of Python from the official website https://www.python.org/downloads/
4. Run the installer and make sure to add Python to the system PATH during the installation process.
5. Open the Command Prompt and run the command `python --version` to check Python is installed correctly.
6. Check pip version by running `pip -V`, if it's not the latest version update pip using `python -m pip install --upgrade pip`
7. Use the command prompt and run `pip install selenium` to install Selenium.
8. Verify Selenium is installed correctly by running `pip show selenium` to check the version
9. Download the appropriate version of Chromedriver from the official website: https://chromedriver.chromium.org/downloads
10. Extract the downloaded file to a directory on your computer and note the location of the executable file.
11. In your Selenium code, specify the path to the Chromedriver executable using the `webdriver.Chrome(executable_path=[path_to_chromedriver])`
12. Create a new Python file in VS Code and start writing your Selenium code.
13. To run the Python code, use the built-in terminal in VS Code and run the command `python [filename.py]` or
    - press `F5` to run the script, or
    - click on the green run button on the top right corner, or 
    - use the command `Ctrl+Shift+B` to run the script


