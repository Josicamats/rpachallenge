from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
import logging
import re


pdf = PDF()
fs = FileSystem()

class PdfActivities:

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)


#Get Name of Investment from pdf
    def extract_investment_pdf(self, directory):
        try:
            rgx = '(?<=Name of this Investment: )[^\.]+'
            text = pdf.get_text_from_pdf(directory,"1")
            matches = re.findall(rgx, text[1])
            for match in matches:
                return match
        except Exception as err:
            self.logger.error("Get PDF Investment Name fails: " + str(err))
            raise SystemError("Get PDF Investment Name fails: " + str(err))

#Get UUI from pdf
    def extract_uii_pdf(self, directory):
        try:
            text = pdf.get_text_from_pdf(directory, "1")
            rgx = '(?<=Unique Investment Identifier \(UII\): )[0-9]+\-?[0-9]+'
            text = pdf.get_text_from_pdf("./output/393-000000086.pdf","1")
            matches = re.findall(rgx, text[1])
            for match in matches:
                return match
        except Exception as err:
            self.logger.error("Get PDF Unique Investment Identifier fails: " + str(err))
            raise SystemError("Get PDF Unique Investment Identifier fails: " + str(err))


#Find files

    def find_files(self, directory, ext):
        try:
            outfiles=[]
            files = fs.list_files_in_directory(directory)
            for f in files:
                if fs.get_file_extension(f).__eq__(ext):
                    outfiles.append(f)
            return outfiles
        except Exception as err:
            self.logger.error("Find files fails: " + str(err))
            raise SystemError("Find files fails: " + str(err))

