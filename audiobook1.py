from win32com.client import Dispatch
import PyPDF2
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)

book = open('oops_tutorial.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(book)
pages = pdfReader.numPages
print(pages)
print(Fore.WHITE+Back.BLACK+"What is Object Oriented Programming? ")

for num in range(9, pages):
    page = pdfReader.getPage(7)
    text = page.extractText()
    def speak(str):
        speak = Dispatch(("SAPI.SpVoice"))
        speak.Speak(str)

    if __name__ == '__main__':
        speak((text))
        print(Fore.BLUE + Back.YELLOW + text)
