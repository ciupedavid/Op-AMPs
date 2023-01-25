# Step-by-step tutorial for installing Python, Selenium and Chromedriver using VS Code

1. Download and install VS Code from the official website: https://code.visualstudio.com/download
2. Open VS Code and go to the Extensions tab. Search for "Python" and install the Python extension by Microsoft. 
3. ![Alternate image text](https://miro.medium.com/max/828/0*I9tSUwyQxYxz1hnM)
4. Download the latest version of Python from the official website https://www.python.org/downloads/
5. Run the installer and make sure to add Python to the system PATH during the installation process.
6. Open the Command Prompt and run the command `python --version` to check Python is installed correctly.
7. Check pip version by running `pip -V`, if it's not the latest version update pip using `python -m pip install --upgrade pip`
8. ![Alternate image text](https://miro.medium.com/max/828/0*-El-bHKVyD8MIGrF)
9. Use the command prompt and run `pip install selenium` to install Selenium.
10. Verify Selenium is installed correctly by running `pip show selenium` to check the version
11. Download the appropriate version of Chromedriver from the official website: https://sites.google.com/a/chromium.org/chromedriver/downloads
![Alternate image text](https://miro.medium.com/max/828/1*52QQNczdZe6jZzn083Q_tg.webp)
12. Extract the downloaded file to a directory on your computer and note the location of the executable file.
13. In your Selenium code, specify the path to the Chromedriver executable using the `webdriver.Chrome(executable_path=[path_to_chromedriver])`
14. Create a new Python file in VS Code and start writing your Selenium code.

Note that, you should check the version of chrome browser and download the same version of chrome driver. Also, make sure that the chrome driver executable is in the system path or you have to provide the full path when initializing the web driver.
