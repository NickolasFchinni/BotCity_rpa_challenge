"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/web/
"""


# Import for the Web Bot
from botcity.web import WebBot, Browser, By

import pandas

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    
    uxPath = "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button"
    InputsDicionary = {
        "inputCompanyRole": '//input[@ng-reflect-name="labelRole"]',
        "inputFirstName": '//input[@ng-reflect-name="labelFirstName"]',
        "inputLastName": '//input[@ng-reflect-name="labelLastName"]',
        "inputAdress": '//input[@ng-reflect-name="labelAddress"]',
        "inputCompanyName": '//input[@ng-reflect-name="labelCompanyName"]',
        "inputEmail": '//input[@ng-reflect-name="labelEmail"]',
        "inputPhone": '//input[@ng-reflect-name="labelPhone"]'
    }
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\Users\nulle\Downloads\crhomedriver\chromedriver.exe"

    # Opens the BotCity website.
    bot.browse("https://rpachallenge.com/")
    bot.maximize_window()
    buttonStart = bot.find_element(uxPath, by=By.XPATH)
    buttonStart.click()
    # Implement here your logic...
    savingExcel = pandas.read_excel(r"D:\PythonProjects\PythonProjectRPA\rpa_challenge\challenge.xlsx")

    for index, row in savingExcel.iterrows():
        bot.find_element(InputsDicionary["inputFirstName"], by=By.XPATH).send_keys(row['First Name'])
        bot.find_element(InputsDicionary["inputLastName"], by=By.XPATH).send_keys(row['Last Name '])
        bot.find_element(InputsDicionary["inputAdress"], by=By.XPATH).send_keys(row['Address'])
        bot.find_element(InputsDicionary["inputCompanyName"], by=By.XPATH).send_keys(row['Company Name'])
        bot.find_element(InputsDicionary["inputEmail"], by=By.XPATH).send_keys(row['Email'])
        bot.find_element(InputsDicionary["inputPhone"], by=By.XPATH).send_keys(row['Phone Number'])
        bot.find_element(InputsDicionary["inputCompanyRole"], by=By.XPATH).send_keys(row['Role in Company'])
        bot.find_element("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input", by=By.XPATH).click()

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
