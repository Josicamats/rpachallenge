import logging
from RPA.Browser.Selenium import Selenium
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem

browser = Selenium()
tb = Tables()
fs = FileSystem()

class WebNavigation:

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)

#set download directory
    def set_download_directory(self, directory):
        try:
            browser.set_download_directory(directory, True)
        except Exception as err:
            self.logger.error("Set download directory fails: " + str(err))
            raise SystemError("Set download directory fails: " + str(err))

            
# Open specified website 
    def open_website(self, url: str):
        try:
            browser.open_available_browser(url)
            browser.maximize_browser_window()
            browser.set_browser_implicit_wait(30)
            #self.browser.wait_until_page_contains_element()
        except Exception as err:
            self.logger.error("Login website fails bc: " + str(err))
            raise SystemError("Login website fails bc: " + str(err))

# Click specified button
    def click_button(self, button: str):
        try:
            browser.wait_until_element_is_visible(button)
            browser.click_element_when_visible(button)
            #browser.set_browser_implicit_wait(30)
        except Exception as err:
            self.logger.error("Click button failes bc: " + str(err))
            raise SystemError("Click button failes bc: " + str(err))

# Close browsers:
    def close_browser(self):
        try:
            browser.close_all_browsers()
        except Exception as err:
            self.logger.error("Close all browsers fails bc: " + str(err))
            raise SystemError("Close all browsers fails bc: " + str(err))

# Get Agencies Data:
    def get_agencies(self):
        try:
            total = 0
            count_agencies = len(browser.find_elements("(//*[contains(@class, 'col-sm-4 text-center noUnderline')])"))
            agencies = []
            for i in range(1,10):
                for j in range (1,4):
                        total = total + 1
                        if total <= count_agencies:
                            agency = browser.find_element("xpath://*[@id='agency-tiles-widget']/div/div["+str(i)+"]/div["+str(j)+"]/div/div/div/div[1]/a/span[1]").text
                            agencies.append(agency)
                #dt_agencies = tb.create_table(agencies)
            return agencies
        except Exception as err:
            self.logger.error("Unable to get Agencies names bc: " + str(err))
            raise SystemError("Unable to get Agencies names bc: " + str(err))

# Get Amounts Data:
    def get_amounts(self):
        try:
            total = 0
            count_agencies = len(browser.find_elements("(//*[contains(@class, 'col-sm-4 text-center noUnderline')])"))
            amounts = []
            for i in range(1,10):
                for j in range (1,4):
                    total = total + 1
                    if total <= count_agencies:
                        amount = browser.find_element("xpath://*[@id='agency-tiles-widget']/div/div["+str(i)+"]/div["+str(j)+"]/div/div/div/div[1]/a/span[2]").text
                        amounts.append(amount)
                #dt_amounts = tb.create_table(amounts)
            return amounts
        except Exception as err:
            self.logger.error("Unable to get amounts of each agency bc: " + str(err))
            raise SystemError("Unable to get amounts of each agency bc: " + str(err))

#Scraping data from table
    def get_table_data(self):
        try:
            browser.select_from_list_by_value("//*[@id='investments-table-object_length']/label/select","-1")
            browser.set_browser_implicit_wait(5)
            row_count = len(browser.find_elements("//*[@id='investments-table-object']/tbody/tr"))
            col_count = len(browser.find_elements("//*[@id='investments-table-object']/tbody/tr[1]/td"))
            data = tb.create_table()
            columns = ["A","B","C","D","E","F","G"]
            for col in columns:
                tb.add_table_column(data, col)
            for n in range(1, row_count+1):
                browser.select_from_list_by_value("//*[@id='investments-table-object_length']/label/select","-1")
                rows = []
                row = 0
                for m in range(1, col_count+1):
                    browser.select_from_list_by_value("//*[@id='investments-table-object_length']/label/select","-1")
                    path = "//*[@id='investments-table-object']/tbody/tr["+str(n)+"]/td["+str(m)+"]"
                    table_data = browser.find_element(path).text
                    rows.append(table_data)
                    if(columns[row] == 'A'):
                        directory = "C:\\ROBOTS\\RPAChallenge\\Challenge\\output"
                        download_pdf(table_data, directory)
                    row = row + 1
                tb.add_table_row(data, rows)
            return data
        except Exception as err:
            self.logger.error("Scraping data from table fails: " + str(err))
            raise SystemError("Scraping data from table fails: " + str(err))

# Download Specified Business Case if Exists Link
def download_pdf(file: str, directory):
    try:
        tableURL = "/drupal/summary/393/" + file
        exist = browser.does_page_contain_link(tableURL)
        if(exist):
            link = browser.find_element('//a[@href="'+tableURL+'"]')
            browser.click_link(link)
            browser.set_browser_implicit_wait(30)
            pdfPath = browser.find_element("//*[@id='business-case-pdf']/a")
            browser.click_link(pdfPath)
            while(fs.does_file_not_exist(directory+"\\"+file+".pdf")):
                browser.set_browser_implicit_wait(10)
            browser.go_back()
            browser.go_back()
    except: 
        pass
    