import requests
from bs4 import BeautifulSoup
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import textwrap

SERVICE_ACCOUNT_FILE = 'C:\\Users\\Etern\\Desktop\\Desktop\\Python\\amazon review scrapper\\amazon-reviews-scraper-ee1ee330d0b7.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE).with_scopes(SCOPES)

def get_amazon_reviews(product_url):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(executable_path='C:\\Users\\Etern\\Desktop\\Desktop\\Python\\amazon review scrapper\\chromedriver\\chromedriver', options=options)
    driver.get(product_url)

    # Navigate to the "See all reviews" page
    see_all_reviews = WebDriverWait(driver, 10000).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-hook='see-all-reviews-link-foot']")))
    driver.get(see_all_reviews.get_attribute('href'))

    reviews = []
    page = 1

    while True:
        try:
            WebDriverWait(driver, 10000).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-hook='review']")))
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            reviews_section = soup.find_all('div', {'data-hook': 'review'})

            if not reviews_section:
                break

            for i, review in enumerate(reviews_section):
                try:
                    username = review.find('span', {'class': 'a-profile-name'}).text
                    stars = review.find('span', {'class': 'a-icon-alt'}).text
                    title = review.find('a', {'data-hook': 'review-title'}).text
                    text = review.find('span', {'data-hook': 'review-body'}).text
                    # Check if text is longer than 50 characters and contains a space
                    if len(text) > 150 and ' ' in text:
                        # Split text into multiple lines
                        text = textwrap.fill(text, width=150)
                    reviews.append([username, stars, title, text])
                    print(f"Review {len(reviews)} added from page {page}: {username}, {stars}, {title}, {text}")
                except Exception as e:
                    print(f"Error encountered while parsing review {i + 1} on page {page}: {str(e)}")

            next_page = driver.find_elements(By.CSS_SELECTOR, 'a[href*="pageNumber"]')

            if not next_page:
                break

            next_page_link = next_page[-1]  # get the last "Next page" link in the list
            next_page_text = next_page_link.text.strip()

            if "Next page" not in next_page_text:
                break

            # Handle StaleElementReferenceException
            while True:
                try:
                    next_page_link.click()
                    break
                except StaleElementReferenceException:
                    next_page_link = driver.find_elements(By.CSS_SELECTOR, 'a[href*="pageNumber"]')[-1]

            page += 1

        except Exception as e:
            print(f"Error encountered: {str(e)}")
            break


    driver.quit()
    return reviews



def write_to_google_sheet(sheet_id, data, headers):

    service = build('sheets', 'v4', credentials=credentials)

    # Set column widths based on the length of the headers and data
    sheet_properties = {
        'requests': [
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': 0,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': len(headers)
                    },
                    'properties': {
                        'pixelSize': max([max([len(str(item)) for item in row]) for row in data]) * 1,

                    },
                    'fields': 'pixelSize'
                }
            }
        ]
    }

    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=sheet_properties).execute()

    try:
        range_name = 'Sheet1!A1:E1'
        body = {
            'values': data
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=sheet_id, range=range_name,
            valueInputOption='RAW', insertDataOption='INSERT_ROWS',
            body=body).execute()
        print(f"{result.get('updates').get('updatedCells')} cells appended.")

    except HttpError as error:
        print(f"An error occurred: {error}")



def create_and_share(sheet_title, email):
    service = build('sheets', 'v4', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials)

    spreadsheet = {
        'properties': {
            'title': sheet_title
        }
    }

    spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
    print(f"Google Sheet created with ID {spreadsheet.get('spreadsheetId')}.")

    file_id = spreadsheet.get('spreadsheetId')
    user_permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': email
    }

    command = drive_service.permissions().create(fileId=file_id, body=user_permission, fields='id')
    command.execute()

    print(f"Google Sheet shared with {email}.")
    sheet_url = f"https://docs.google.com/spreadsheets/d/{file_id}/edit"
    print(f"Sheet URL: {sheet_url}")

    return file_id

if __name__ == "__main__":
    product_url = input("Enter the Amazon product URL: ")
    sheet_title = input("Enter the Google Sheet title: ")
    your_email = input("Enter your email address: ")

    reviews = get_amazon_reviews(product_url)
    print(f"Parsed {len(reviews)} reviews in total.")
    headers = ['Username', 'Stars', 'Review Title', 'Review Text']
    data = [headers] + reviews

    sheet_id = create_and_share(sheet_title, your_email)
    write_to_google_sheet(sheet_id, data, headers)

