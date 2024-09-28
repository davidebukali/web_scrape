import random
import requests
import requests_cache
import openpyxl
import re
import time
from bs4 import BeautifulSoup

# Cache requests to avoid hitting the server multiple times for the same URL
requests_cache.install_cache('yellow_ug_cache', expire_after=86400)  # Cache for 24 hours
user_agents = [
    "(Windows NT 10.0; Win64; x64) Gecko/20100101 Firefox/89.0",
    "(Macintosh; Intel Mac OS X 10_15_7) Safari/605.1.15",
    "(Linux; Android 10; Pixel 3 XL Build/QP1A.190711.020; wv) Chrome/83.0.4103.106",
    "(iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Safari/604.1",
    "(Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107",
    "(X11; Ubuntu; Linux x86_64; rv:88.0) Gecko/20100101 Firefox/88.0",
    "Mozilla/5.0 (Linux; Android 9; SM-G960F Build/PPR1.180610.011; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0",
]

# Step 2: Generate a random user agent
random_user_agent = random.choice(user_agents)

# Step 3: Make a request using the random user agent
headers = {
    'User-Agent': random_user_agent
}

file_path = "yellowug.xlsx"
# Create a new Excel workbook
workbook = openpyxl.load_workbook(file_path)

# Select the active sheet
sheet = workbook.active

root_url = "https://www.yellow.ug"
business_directory_url = "https://www.yellow.ug/browse-business-directory"

def remove_illegal_characters(text):
    # Regular expression to match illegal characters (control characters and NULL byte)
    return re.sub(r'[\x00-\x1F]', '', text)

def get_last_page(pages_container):
    if pages_container:
    # Get all the li elements within the ul
        pagination_list = pages_container.find('ul').find_all('li')

        # Skip the first (back button) and last (forward button) li elements
        page_numbers = []

        for li in pagination_list:  # Slicing to exclude first and last
            li_text = li.get_text()
            if li_text.isdigit():
                page_numbers.append(li.get_text())  # Append the page number

        # Get the last page number from the extracted numbers
        last_page = page_numbers[-1] if len(page_numbers) > 0 else 1

        return int(last_page)
    else:
        return 1

def scrape_company_detail(url):
    # Fetch the content of the web page
    response = requests.get(url, headers=headers)

    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')
    company_detail = []
    section = soup.find('section')
    if section:
        h1 = section.find('h1')
        important_div = section.find('div', class_='important_tag')
        description = section.find('div', class_='desc')
        phone = section.find('div', class_='phone')
        mobile_label = section.find('div', class_='label', string=lambda text: text and 'Mobile phone' in text)
        mobile_div_text = ""
        website = section.find('div', class_='weblinks')
        location = section.find('div', class_='location')
        map = section.find('div', id='map_canvas')
        tags = section.find('div', class_='tags')
        
        if h1:
            company_detail.append(h1.get_text())
        
        # Get the text content if the div is found
        if important_div:
            important_text = important_div.get_text()
            company_detail.append(important_text)
        else:
            company_detail.append("")

        if description:
            description_text = description.get_text()
            company_detail.append(remove_illegal_characters(description_text))
        else:
            company_detail.append("")
        
        if mobile_label:
            mobile_div = mobile_label.find_next_sibling()
            mobile_div_text = mobile_div.get_text()

        if phone:
            phone_text = phone.get_text()
            company_detail.append(f"{phone_text}, {mobile_div_text}")
        else:            
            company_detail.append(f"{mobile_div_text}")

        if website:
            website_text = website.get_text()
            company_detail.append(website_text)
        else:            
            company_detail.append("")

        if location:
            location_text = ''.join([str(content) for content in location.contents if isinstance(content, str) or (content.name == 'a')]).strip()
            company_detail.append(location_text)
        else:            
            company_detail.append("")
            
        if map:
            data_map_ltd_value = map.get('data-map-ltd')
            data_map_lng_value = map.get('data-map-lng')
            company_detail.append(f"{data_map_ltd_value},{data_map_lng_value}")
        else:            
            company_detail.append("")

        if tags:
            tag_text = ', '.join([a.get_text() for a in tags.find_all('a') if not a.has_attr('class')]).strip()
            company_detail.append(tag_text)
        else:            
            company_detail.append("")
            
    sheet.append(company_detail)

def traverse_company_list(url):
    page_number = 1

    # Loop through all pages
    while True:
        # Fetch the content of the web page
        response = requests.get(f"{root_url}{url}/{page_number}", headers=headers)

        # Parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')

        pages_container = soup.find('div', class_='pages_container')

        last_page_number = get_last_page(pages_container)

        if page_number > last_page_number:
            print(f"Breaking at Last page {last_page_number}.")
            break

        # Find all divs with id that contains 'listings'
        company_parent = soup.find('div', id=lambda x: x and 'listings' in x)

        company_divs = company_parent.find_all('div', class_='company')
        if not company_divs:
            print(f"No companies found on page {page_number}. Ending pagination.")
            break

        # Extract href from the anchor tags within those divs
        for div in company_divs:
            h4 = div.find('h4')
            if h4:
                print(h4.get_text())
                a_tag = h4.find('a')  # Find anchor inside h4
                if a_tag:
                    href = a_tag.get('href')  # Get the href attribute
                    scrape_company_detail(f"{root_url}{href}")
        
        print(f"Done scraping page {page_number}.")
        page_number += 1 
        # Delay to avoid overwhelming the server
        print(f"Delay to avoid overwhelming the server {page_number}.")
        delay = random.uniform(2, 5)
        time.sleep(delay)  # 2 seconds delay between requests
        # Save the workbook
        workbook.save("yellowug.xlsx")

def traverse_company_categories():
    # Fetch the content of the web page
    response = requests.get(business_directory_url, headers=headers)

    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find all divs with class that contains 'cmap'
    icats = soup.find_all('ul', class_='icats')

    # Extract href from the anchor tags within those divs
    li_items = icats[1].find_all('li')
    sliced_li_item = li_items[4]
    anchor = sliced_li_item.find('a')
    if anchor:
        delay = random.uniform(2, 6)
        print(f"Category {anchor.text}.")
        print(f"Delay to avoid overwhelming the server {delay}.")
        time.sleep(delay)  # 2 seconds delay between requests
        traverse_company_list(anchor['href'])

traverse_company_categories()

# Close the workbook when done
workbook.close()

print("Success")
