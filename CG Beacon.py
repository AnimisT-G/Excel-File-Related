from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from os import listdir
from pathlib import Path
from time import sleep
import openpyxl as xl

column_filter_flag = True


class BROWSER():
    def __init__(self):  # Opening Automated Webpage in Browser
        self.browser = webdriver.Edge()
        self.browser.get("https://cap.brightscope.com/search/beacon/#/plan")

    def sign_in(self, email, password):  # Sign In into Beacon
        self.browser.find_element(By.NAME, "email").click()
        action = ActionChains(self.browser)
        action.key_down(Keys.CONTROL).send_keys(
            'A').key_up(Keys.CONTROL).perform()
        self.browser.find_element(By.NAME, "email").send_keys(email)
        self.browser.find_element(By.NAME, "password").send_keys(password)
        # self.browser.find_element(By.TAG_NAME, "button").click()
        while True:
            try:
                self.browser.find_element(
                    By.CSS_SELECTOR, "button.main-search-button").click()
                return None
            except (NoSuchElementException, ElementNotInteractableException):
                try:
                    self.browser.find_element(By.CSS_SELECTOR, "ul.errors")
                    email, password, removals = email_n_password()
                    self.sign_in(email, password)
                    return None
                except NoSuchElementException:
                    pass

    def search_filters(self):  # Activate Plan Name in Search Filter Navigation
        flag = False
        plan_name_checkbox = self.browser.find_element(
            By.XPATH, '//*[@id="filter-nav"]/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[14]/bale-filter/div/div[1]/md-checkbox/div[2]')
        while not flag:
            flag = self.browser.find_element(
                By.ID, "filter-nav").is_displayed()
            if flag:
                self.browser.find_elements(By.TAG_NAME, "v-pane")[1].click()
                flag = False
                while not flag:
                    flag = plan_name_checkbox.is_displayed()
        plan_name_checkbox.click()

    def column_filters(self):  # Activate EIN & Company Name in Column Filter Navigation
        self.browser.find_element(By.CSS_SELECTOR, "div.col-button").click()
        column_filter_navigation = self.browser.find_element(
            By.CSS_SELECTOR, "md-sidenav.md-sidenav-right")
        flag = False
        while not flag:
            flag = column_filter_navigation.is_displayed()
            if flag:
                self.browser.find_element(
                    By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]').click()
                flag = False
                while not flag:
                    flag = self.browser.find_element(
                        By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[11]/div/md-checkbox').is_displayed()
        self.browser.find_element(
            By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[11]/div/md-checkbox/div[2]').click()
        self.browser.find_element(
            By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[13]/div/md-checkbox/div[2]').click()
        send_us_email = self.browser.find_element(
            By.ID, "hbl-live-chat-wrapper")
        self.browser.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, send_us_email)
        self.browser.find_element(
            By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[3]/button/div[1]').click()
        while column_filter_navigation.is_displayed():
            pass

    def plan_name_search(self, pname):  # Type Plan Name in Field and Run Search
        search_filter_navigation = self.browser.find_element(
            By.ID, "filter-nav")
        flag = search_filter_navigation.is_displayed()
        if not flag:
            try:
                self.browser.find_element(
                    By.CSS_SELECTOR, "div.company-information").click()
            except (NoSuchElementException, ElementNotInteractableException):
                self.browser.find_element(
                    By.CSS_SELECTOR, "button.main-search-button").click()
        while True:
            try:
                self.browser.find_element(By.ID, "fl-input-125").click()
                break
            except (NoSuchElementException, ElementNotInteractableException):
                pass
        plan_name_input_field = self.browser.find_element(
            By.ID, "fl-input-125")
        action = ActionChains(self.browser)
        action.key_down(Keys.CONTROL).send_keys(
            'A').key_up(Keys.CONTROL).perform()
        plan_name_input_field.send_keys(pname)
        flag = False
        plan_name_field_list = self.browser.find_element(By.ID, "ul-125")
        sleep(1)
        while True:
            try:
                sleep(1)
                flag = plan_name_field_list.is_displayed()
                if flag:
                    plan_name_field_list.click()
                    break
                else:
                    plan_name_input_field.click()
            except (NoSuchElementException, ElementNotInteractableException):
                pass
        self.browser.find_element(
            By.CSS_SELECTOR, "button.submit-button").click()
        flag = True
        while flag:
            flag = self.browser.find_element(
                By.CSS_SELECTOR, "md-sidenav.md-sidenav-right").is_displayed()

    def results(self):  # Search Results
        flag = False
        while not flag:
            try:
                flag = self.browser.find_element(
                    By.CSS_SELECTOR, "span.small-header").is_displayed()
                global column_filter_flag
                if column_filter_flag:
                    self.column_filters()
                    column_filter_flag = False
            except NoSuchElementException:
                try:
                    self.browser.find_element(
                        By.CSS_SELECTOR, "button.main-search-button").click()
                    return 0, self.browser.find_elements(By.CSS_SELECTOR, "div.search-apology")
                except NoSuchElementException:
                    pass
        search_results = self.browser.find_element(
            By.CSS_SELECTOR, "span.small-header").text
        results_plan = self.browser.find_elements(
            By.CSS_SELECTOR, "div.result-cell")
        search_results = search_results[16:-1]
        if len(search_results) > 2:
            search_results = 30
        return int(search_results), results_plan

    def quit(self):  # Close the Automated Webpage
        self.browser.quit()


def main():  # Main Function
    # Email ID and Password for LogIn into Beacon
    try:
        with open("login.txt", 'r') as file:
            lines = file.readlines()
            email, password, removals = lines[0], lines[1], lines[2].split(':')
    except FileNotFoundError:
        email, password, removals = email_n_password()

    file_name = excel_file_name_input()
    print(f'File - {file_name}\nOpening File...')
    data_excel = xl.load_workbook(Path(file_name))
    data_sheet = data_excel.active
    new_excel = xl.Workbook()
    sheet = new_excel.active
    for row in range(1, data_sheet.max_row):
        sheet[f"A{row}"].value = data_sheet[f"B{row}"].value
        sheet[f"B{row}"].value = data_sheet[f"Y{row}"].value
        sheet[f"C{row}"].value = data_sheet[f"AP{row}"].value

    data_excel.close()

    # Creating the Web Automation Object
    driver = BROWSER()
    driver.sign_in(email, password)
    driver.search_filters()
    previous_plan, previous_c_plan = 'a'*2
    for row in range(2, sheet.max_row):
        plan = sheet[f"B{row}"].value
        if sheet[f"C{row}"].value != None:
            continue
        c_plan = check_plan(plan, removals)
        if plan in [previous_plan, 'NULL', None] or c_plan == previous_c_plan:
            sheet[f"D{row}"].value = sheet[f"D{row - 1}"].value
            continue
        previous_plan = plan
        previous_c_plan = c_plan
        driver.plan_name_search(c_plan)
        search_results, results_plan = driver.results()
        if search_results == 0:
            sheet[f"D{row}"].value = 'Not Found'
            continue
        elif search_results > 29:
            continue

        Results = []
        for i in range(5):
            Results.append([results_plan.pop(0).text])
        sheet[f"D{row}"].value = 'Found'
        print(Results)
        for i in results_plan:
            print(i.text)
    driver.quit()

    new_excel.save(f"./Automated Results.xlsx")
    new_excel.close()


def email_n_password():  # Asking User Email & Password if Unable to Login
    email = input("Enter Login Email    : ")
    password = input("Enter Login Password : ")
    removals = "inc:llc:pc:p.c:pllc:cpas:,inc:,llc:plan:and:trust:prof"
    option = input("\nSAVE Login Credentials in 'login.txt' file (Y) : ")
    if option in ('Y', 'y'):
        with open("login.txt", 'w') as file:
            file.write(email+'\n'+password+'\n'+removals)
    return email, password, removals.split(':')


def excel_file_name_input():
    print("Excel files in same directory:")
    files = listdir('.')
    xl_files = []
    counter = 0
    for file in files:
        if file[-4:] == 'xlsx':
            counter += 1
            xl_files.append(file)
            print(f"{counter}. {file}")
    if counter == 1:
        return xl_files[0]
    elif counter == 0:
        input("\n**No Excel Files Found.\n  Closing Program !!")
        exit()
    option = int(
        input(f"\nEnter the file number corresponding to file name: ")) - 1
    if len(xl_files) > option >= 0:
        return xl_files[option]
    else:
        input("\n\n**Wrong Input !!!\n  Closing Program !!")
        exit()


def check_plan(plan='', removals=[]):
    temp = plan.lower().split(' ')
    removing_indexes = []
    for i in temp:
        if i in removals:
            removing_indexes.append(temp.index(i))
    removing_indexes.reverse()
    for i in removing_indexes:
        temp.pop(i)

    plan = ''
    for i in temp:
        c = i.find(',')
        if c > 0:
            i = i[:c] + i[c+1:]
        c = i.find("'s")
        if c > 0:
            i = i[:-2]
        c = i.find(".")
        if c > 0:
            i = i[:c] + ' ' + i[c+1:]
        plan += i + ' '
    return plan


main()
