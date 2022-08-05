import requests
from bs4 import BeautifulSoup
import xlwings as xw
from loguru import logger


def main(domain, endpoint):
    data_count = 0
    logger.info("Create workbook")
    wb = xw.Book('Opencorp-companies.xlsx')
    worksheet = wb.sheets('Companies')

    logger.info("Starting scraper!!!")
    response = requests.get(domain+endpoint)
    soup = BeautifulSoup(response.text, "html.parser")

    for a in soup.find_all('a', 'add_filter'):
        logger.info(f"Requesting to {domain+a.attrs['href']}")
        response = requests.get(domain+a.attrs['href']) 
        soup = BeautifulSoup(response.text, "html.parser")

        for a in soup.find_all('a', 'company_search_result'):
            data_count += 1
            logger.debug(f"Writing {a.contents[0]}{domain+a.attrs['href']}...")
            worksheet.range(f'A{data_count}').value = a.contents[0]
            worksheet.range(f'B{data_count}').value = domain+a.attrs['href']


if __name__ == "__main__":
    domain = "https://opencorporates.com"
    endpoint = "/companies?q=A&utf8"

    main(domain, endpoint)







    


    






