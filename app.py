from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import time
import requests
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#-----------------------------------------------------------------#

# Create a workbook and a worksheet to store the event details
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Event Details"

# Write headers to the Excel file
ws.append(["Name", "Event URL", "Date & Time", "Location", "Host Name", "Event Tags"])

#-----------------------------------------------------------------#

# URL of the website listing popular Toronto restaurants or events
url = "https://www.meetup.com/find/?keywords=restaurant&location=ca--on--Toronto&source=EVENTS"

# Set up the Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--ignore-ssl-errors")

driver = webdriver.Chrome(options=chrome_options)
driver.get(url)
time.sleep(3)  # Give the page time to load initially

# Initialize variables for scrolling
scroll_pause_time = 3
max_scroll_attempts = 20
new_events_loaded = True
attempts = 0

# Initialize restaurants as an empty list
restaurants = []

# Set up scrolling to load all records dynamically
while new_events_loaded and attempts < max_scroll_attempts:
    # Scroll to the bottom of the page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(scroll_pause_time)

    # Parse page content after each scroll
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    new_restaurants = soup.find_all(
        'div', 
        class_='relative z-0 flex h-full break-words bg-transparent bg-white bg-cover bg-clip-padding p-0 transition-shadow duration-300 w-full flex-row justify-start py-4 border-t border-gray3 md:pt-4 md:pb-5'
    )
    
    # Append only new restaurants to avoid duplicates
    for new_restaurant in new_restaurants:
        if new_restaurant not in restaurants:
            restaurants.append(new_restaurant)

    # Update attempts count based on newly loaded events
    if len(restaurants) > attempts * 20:  # Assuming 20 events load per scroll
        attempts += 1
    else:
        new_events_loaded = False

driver.quit()  # Close the driver after scraping

# Loop through each restaurant/event and extract relevant details
for restaurant in restaurants:
    # Extract name
    name_tag = restaurant.find('h2', class_='text-gray7 font-medium text-base pb-1 pt-0 line-clamp-3')
    name = name_tag.text.strip() if name_tag else "N/A"
    
    # Extract event URL
    event_url_tag = restaurant.find('a', class_='w-full cursor-pointer hover:no-underline')
    event_url = event_url_tag['href'] if event_url_tag else "URL not available"

    # Fetch detailed event page
    eventResponse = requests.get(event_url)
    event_soup = BeautifulSoup(eventResponse.text, 'html.parser')
    
    # Extract date and time
    time_tag = restaurant.find('time', {'datetime': True})
    date_time_text = time_tag.get_text(" ", strip=True) if time_tag else "No date available"

    # Extract location
    location_tag = event_soup.find('div', class_='text-gray6', attrs={'data-testid': 'location-info'})
    location = location_tag.text.strip() if location_tag else "Location not available"

    # Extract host name
    host_name_tag = event_soup.find('div', class_='ml-6')
    host_name = "Host Name not available"
    if host_name_tag:
        host_name_tag = host_name_tag.find('span', class_='font-medium')
        host_name = host_name_tag.text.strip() if host_name_tag else "Host Name not available"

    # Extract tags
    tags = [a.get_text() for a in event_soup.find_all('a', class_='tag--topic')]

    # Export data to Excel
    ws.append([name, event_url, date_time_text, location, host_name, ", ".join(tags)])

# Save the Excel file
wb.save("C:/Users/DELL/OneDrive - Cambrian College/College/Capstone Project/Event_Details.xlsx")
