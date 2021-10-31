import logging
import sys
import os
from WebNavigation import WebNavigation
from ExcelActivities import ExcelActivities
from SendEmail import SendEmail
from PdfActivities import PdfActivities

navigation = WebNavigation()
excel = ExcelActivities()
email = SendEmail()
pdf = PdfActivities()
REPORT_PATH = os.environ["REPORT_PATH"]
OUTPUT_PATH = os.environ["OUTPUT_PATH"]
REPORT_NAME = os.environ["REPORT_NAME"]
RECIPIENTS = os.environ["RECIPIENTS"]
SUBJECT = os.environ["SUBJECT"]
URL = os.environ["URL"]
DIVE_IN_BUTTON = os.environ["DIVE_IN_BUTTON"]
AGENCY_BUTTON = os.environ["AGENCY_BUTTON"]

stdout = logging.StreamHandler(sys.stdout)
logging.basicConfig(
    level=logging.INFO,
    format="[{%(filename)s:%(lineno)d} %(levelname)s - %(message)s",
    handlers=[stdout],
)

LOGGER = logging.getLogger(__name__)

def main():
    try:
        logging.info("Process Start..")
        navigation.set_download_directory("C:\\ROBOTS\\RPAChallenge\\Challenge\\output")
        navigation.open_website(URL)
        navigation.click_button(DIVE_IN_BUTTON)
        agencies = navigation.get_agencies()
        amounts = navigation.get_amounts()
        excel.create_file(REPORT_NAME, OUTPUT_PATH)
        dt_agencies = excel.create_table(agencies, amounts)
        excel.write_to_wb('Agencies',REPORT_PATH, dt_agencies)
        navigation.click_button(AGENCY_BUTTON)
        dt_investments = navigation.get_table_data()
        excel.write_to_wb('Investments',REPORT_PATH, dt_investments)
        files = pdf.find_files(OUTPUT_PATH, ".pdf")
        for f in files:
            investmentName = pdf.extract_investment_pdf(f)
            uiiNumber = pdf.extract_uii_pdf(f)
            excel.compare_value(REPORT_PATH, 'Investments', "Investment Title", investmentName)
            print("The value "+ str(investmentName) +" is in the file")
            excel.compare_value(REPORT_PATH, 'Investments', "UII", uiiNumber)
            print("The value "+ str(uiiNumber) +" is in the file")

    except Exception as err:
        logging.info("Sending Error Email")
        email.send_email(RECIPIENTS,SUBJECT,"Process failed: \n " + str(err))
    finally:
        logging.info("Closing Browsers")
        navigation.close_browser()
        logging.info("Sending Completed Email")
        email.send_email(RECIPIENTS,SUBJECT,'Process finished.')


if __name__ == "__main__":
    
    main()
