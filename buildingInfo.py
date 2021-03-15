from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import numpy as np
import csv

my_data = np.genfromtxt('Output.csv', dtype=str, delimiter=',', skip_header=1)

my_file = csv.writer(open('buildingInfo3.csv', 'w', newline=''))
my_file.writerow(['Constituent ID', 'Real Estate Number', 'Year Built', 'Heated Area (sq ft)', 'Stories', 'Bedrooms', 'Baths', 'Rooms / Units'])

for i, row in enumerate(my_data):
    rowValue = [row[0], row[1]]     # add Constituent ID + Real Estate Number

    if row[2] == 'n/a':             # if no RE# provided
        my_file.writerow(rowValue)
        continue
    else:                           # if RE# is provided
        browser = webdriver.Chrome('C:/WebDriver/bin/chromedriver')
        browser.get("https://paopropertysearch.coj.net/Basic/Search.aspx")

        re1_search_input = browser.find_element_by_id('ctl00_cphBody_tbRE6')
        re1_search_input.send_keys(row[2])       # row[2] = RE1
        re2_search_input = browser.find_element_by_id('ctl00_cphBody_tbRE4')
        re2_search_input.send_keys(row[3])       # row[3] = RE2

        search_button = browser.find_element_by_id('ctl00_cphBody_bSearch')
        search_button.click()

        try:
            text = "Detail.aspx"
            property_link = browser.find_element_by_xpath('//a[contains(@href, "%s")]' % text)
            property_link.click()
        except NoSuchElementException:
            browser.close()
            my_file.writerow(rowValue)
            continue

        try:
            yearBuilt = browser.find_element_by_id('ctl00_cphBody_repeaterBuilding_ctl00_lblYearBuilt')
            rowValue.append(yearBuilt.text)
        except NoSuchElementException:
            rowValue.append('')

        for i in range(2, 9):
            xpath = '//*[@id="ctl00_cphBody_repeaterBuilding_ctl00_gridBuildingArea"]/tbody/tr[{}]/td[1]'.format(i)
            try:
                currentElement = browser.find_element_by_xpath(xpath)
                if currentElement.text == "Total":
                    xpath = '//*[@id="ctl00_cphBody_repeaterBuilding_ctl00_gridBuildingArea"]/tbody/tr[{}]/td[3]'.format(i)
                    heatedArea = browser.find_element_by_xpath(xpath)
                    rowValue.append(heatedArea.text)
                    break
            except (NoSuchElementException, IndexError):
                rowValue.append('')
        try:
            if heatedArea is None:
                rowValue.append('')
        except NoSuchElementException:
            rowValue.append('')

        try:
            stories = browser.find_elements_by_class_name('col_code')[9]
            rowValue.append(stories.text)
        except (NoSuchElementException, IndexError):
            rowValue.append('')

        try:
            bedrooms = browser.find_elements_by_class_name('col_code')[10]
            rowValue.append(bedrooms.text)
        except (NoSuchElementException, IndexError):
            rowValue.append('')

        try:
            baths = browser.find_elements_by_class_name('col_code')[11]
            rowValue.append(baths.text)
        except (NoSuchElementException, IndexError):
            rowValue.append('')

        try:
            rooms = browser.find_elements_by_class_name('col_code')[12]
            rowValue.append(rooms.text)
        except (NoSuchElementException, IndexError):
            rowValue.append('')

        browser.close()
        my_file.writerow(rowValue)

my_file.close()
