# imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from win32com.client import Dispatch
import time
from datetime import date
import sys
import pyautogui

'''
op = webDriver.ChromeOptions()
op.add_argument("headless")
'''


# Speech function
def speak_function(sentence):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(sentence)


# func to find ERROR MESSAGES elements in the web page
def find_element(id_para):
    try:
        driver.find_element(By.ID, id_para)
    except NoSuchElementException:
        return False
    return True


# func try again
def try_again():
    speak_function('click 1 to try again')
    again = input("Click any button to close\nClick '1' to try again\n")
    if again == '1':
        exec(open("bookGas.py").read())
    else:
        exit(0)


print("\n")

# load chrome driver
driverLocation = 'chromedriver.exe'
driver = webdriver.Chrome(driverLocation)

# maximize the window
driver.maximize_window()

# open website
driver.get('https://myhpgas.in/myHPGas/HPGas/BookRefill.aspx')

# click on booking link
time.sleep(3)
bookingLink = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[2]/div[2]/table/tbody/tr/td[2]/div/table/tbody/tr/td[2]/b/i/a')
bookingLink.click()

# fill agency name
time.sleep(3)
agencyName = driver.find_element(By.ID, 'ContentPlaceHolder1_txtDistributorName')
agencyName.click()
agencyName.send_keys('SOUNDARYA AGENCIES BANGALORE CENTRAL (14755700)')

# fill customer ID
time.sleep(1)
consumerNumber = driver.find_element(By.ID, 'ContentPlaceHolder1_txtConsumerNoQuick')
consumerNumber.click()
consumerNumber.send_keys('639937')

# scroll to the bottom of the page
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# ask user to type captcha
speak_function('enter the captcha code in 10 seconds cause i am a robot.')

time.sleep(0.5)
captcha = driver.find_element(By.ID, 'ContentPlaceHolder1_MyCaptcha1_tbCaptchaInput')
captcha.click()

# click on submit button
time.sleep(10)
submitButton = driver.find_element(By.ID, 'ContentPlaceHolder1_btnProceed')
submitButton.click()

# if timed out wrt captcha
time.sleep(3)
if find_element("ContentPlaceHolder1_lblErrorMessage"):
    speak_function('you are timed out due to incorrect code')
    print('YOU ARE TIMED OUT BECAUSE CODE WAS NOT TYPED IN PROPERLY\n')
    driver.close()
    try_again()

# if timed out because of previous login
elif find_element('ContentPlaceHolder1_divMain'):
    speak_function('sorry! your previous login session is not ended yet. Please try again after a while')
    print('TRY AGAIN AFTER SOME TIME (~30MIN)\n')
    driver.close()
    sys.exit()

# if neither timed out wrt captcha nor because of previous login
else:
    # fill mobile number
    time.sleep(1)
    mobileNumber = driver.find_element(By.ID, 'ContentPlaceHolder1_txtMobileNo')
    mobileNumber.click()
    mobileNumber.send_keys('9611252969')

    # print the price
    price = driver.find_element(By.ID, 'ContentPlaceHolder1_lblRsp')
    rate = price.get_attribute('innerHTML')
    print('PRICE IS ', rate)
    speak_function(f'price is {rate.split(".")[0]}')

    # ask user if to pay now or not
    speak_function("do you want to pay now? click 1 for yes")
    pay = input("\nPay Now?\nClick '1' for 'YES'\nClick '0' for 'NO'\n")

    # if user wants to pay now
    if pay == '1':
        # click on pay button
        payButton = driver.find_element(By.ID, 'ContentPlaceHolder1_btnOnline')
        payButton.click()

        # if timed out because of previous login
        if find_element('ContentPlaceHolder1_divSuccessMsg'):
            print('\nTRY AGAIN AFTER SOME TIME (~30MIN)\n')
            speak_function('TRY AGAIN AFTER SOME TIME')
            driver.close()

        # if not timed out because of previous login
        else:
            # scroll to the bottom of the page
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # check the T&C box
            time.sleep(0.6)
            termsCheckBox = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$chkAgree')
            termsCheckBox.click()

            # click on accept
            time.sleep(1)
            acceptButton = driver.find_element(By.ID, 'ContentPlaceHolder1_btnAccept')
            acceptButton.click()

            # click on proceed to pay
            time.sleep(2)
            proceedButton = driver.find_element(By.XPATH, '/html/body/center/form/div/div[2]/input')
            proceedButton.click()
            speak_function('Please pay now with a payment method of your choice')

    # if user does not want to pay now
    else:
        speak_function('type one if you wanna exit!\n')
        ext = input('\nExit? ')
        if ext == '1':
            driver.close()
            speak_function('exiting. see you soon. thank you')
            sys.exit(0)
        try_again()


# Success message
while True:
    if find_element('ContentPlaceHolder1_lblSuccessMsg'):
        print("-------------------------\n")
        print('GAS REFILL BOOKED SUCCESSFULLY\n')

        # take a screenshot of the bill
        ss = pyautogui.screenshot()
        ss.save(rf"C:\Users\prakyath.arya\Downloads\LPG-{date.today()}.jpeg")

        speak_function('gas refill booked successfully. thank you and see you soon\n')
        break
exit(0)
