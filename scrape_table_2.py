# %%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pyautogui as pyg
from time import sleep
import pandas as pd

# %%
def open_powerbi():
    print('Opening Power Bi')
    global driver
    power_bi = 'https://app.powerbi.com/view?r=eyJrIjoiNDU4Y2UxNmEtZjc0Yi00ZTkyLTk3N2EtZTEyZTI5MjdkNzQ2IiwidCI6ImI2N2FmMjNmLWMzZjMtNGQzNS04MGM3LWI3MDg1ZjVlZGQ4MSJ9&pageName=ReportSection%20Power%20BI%20Report%20Report%20powered%20by%20Power%20BI'
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(power_bi)
    sleep(3.5)


def text_xpath(xpath: str):
    return str(driver.find_element(By.XPATH, xpath).text)

def text_selector(selector: str):
    return driver.find_element(By.CSS_SELECTOR, selector).text

def click_selector(selector: str):
    driver.find_element(By.CSS_SELECTOR, selector).click()

def click_xpath(xpath: str):
    driver.find_element(By.XPATH, xpath).click()


def expand_table2():
    # Click Nutrient/Bioactive Substance/Enzyme
    click_xpath('/html/body/div[1]/report-embed/div/div/div[1]/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[9]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div[1]/div[1]/div/div/div/div[2]')
    sleep(1.5)
    # Click Focus Mode
    click_xpath('/html/body/div[1]/report-embed/div/div/div[1]/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[9]/transform/div/visual-container-header/div/div/div/visual-header-item-container/div/button')
    sleep(1.5)


def set_columns():
    columns = {}
    position = 2
    while True:
        try:
            xpath = f'/html/body/div[1]/report-embed/div/div/div[1]/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[9]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div[1]/div[1]/div/div/div/div[{position}]'
            columns[text_xpath(xpath).strip()] = position
            position += 1
        except:
            return columns

# %%
def set_first_last_data():
    first_selector = '#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToScreen > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(9) > transform > div > div.visualContent > div > div > visual-modern > div > div > div.tableEx > div.interactive-grid.innerContainer > div.mid-viewport > div > div:nth-child(1)'
    first = int(driver.find_element(By.CSS_SELECTOR, first_selector).get_attribute('row-index')) # determine first visible row

    number_last_selector = 20
    while True: # determine last visible row dynamically
        try:
            last_selector =  f'#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToScreen > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(9) > transform > div > div.visualContent > div > div > visual-modern > div > div > div.tableEx > div.interactive-grid.innerContainer > div.mid-viewport > div > div:nth-child({number_last_selector})'
            last = int(driver.find_element(By.CSS_SELECTOR, last_selector).get_attribute('row-index'))
            break
        except:
            number_last_selector -= 1
    print(f'Scrapping data from {first} to {last}, {number_last_selector} rows')
    return first_selector, first, last, number_last_selector


def create_rows(first, last, columns):
    for row_index in range(last - first + 1): # loop visible range

        linha = {}
        for col, pos in columns.items():
            current_selector = f'#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToScreen > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(9) > transform > div > div.visualContent > div > div > visual-modern > div > div > div.tableEx > div.interactive-grid.innerContainer > div.mid-viewport > div > div:nth-child({row_index+1}) > div > div > div:nth-child({pos})'
            linha[col] = text_selector(current_selector)

        dict_list.append(linha)


def scroll_page(first_selector, last, number_last_selector):
    if number_last_selector == 20:
        last = last + 1 # removed the +1, taking the last row in the previous loop as a margin of error, removing duplicated later in the df
        print(f'Scrolling to {last}')
        while True: # If isn't in the last iteration, it will continue to go down in the BI window
            try:
                current_first = int(driver.find_element(By.CSS_SELECTOR, first_selector).get_attribute('row-index'))
            except:
                print('Did not work, but will try again')
                sleep(0.3)
                continue

            if current_first == last:
                print(f'Found {current_first}')

            if current_first > last:
                raise ValueError('Scroll exceeded limit, script stopped.')

            if (last - current_first) > 9:
                body.send_keys(Keys.PAGE_DOWN)
                sleep(0.22)
            else:
                body.send_keys(Keys.DOWN)
                sleep(0.13)

            try:
                driver.find_element(By.CSS_SELECTOR, '#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToScreen > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(9) > transform > div > div.visualContent > div > div > visual-modern > div > div > div.tableEx > div.interactive-grid.innerContainer > div.mid-viewport > div > div:nth-child(20)')
            except:
                return True
    else:
        print('Scrapping finished!')
        return False


def save_df():
    df = pd.DataFrame(dict_list)
    df = df.drop_duplicates()
    df.to_excel(r'\projects\Power BI\Anvisa_PowerBI_table_2.xlsx', index=False)

# %%
open_powerbi()
expand_table2()
col = set_columns()

# %%
dict_list = []
body = driver.find_element(By.TAG_NAME, "body")
pyg.click(1730, 570)

main_loop_bool = True
while main_loop_bool:
    fs, f, l, nls = set_first_last_data()
    create_rows(f, l, col)
    main_loop_bool = scroll_page(fs, l, nls)
    print('')

save_df()