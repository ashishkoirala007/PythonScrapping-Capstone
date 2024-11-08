import requests
from bs4 import BeautifulSoup
import openpyxl
import time

#-----------------------------------------------------------------#
# Create a workbook and a worksheet to store the event details
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Event Details"

# Write headers to the Excel file
ws.append(["Name", "Event URL", "Date & Time", "Location", "Host Name", "Event Tags"])
#-----------------------------------------------------------------#

# URL of the website listing popular Toronto restaurants or events
url = "https://www.meetup.com/find/?location=ca--on--Toronto&source=EVENTS&keywords=Restaurants"

# Send a request to the website
response = requests.get(url)

# Check if the request was successful
if response.status_code == 200:
    # Parse the HTML content of the page
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find the elements containing the restaurant or event names and details
    restaurants = soup.find_all('div', class_='relative z-0 flex h-full break-words bg-transparent bg-white bg-cover bg-clip-padding p-0 transition-shadow duration-300 w-full flex-row justify-start py-4 border-t border-gray3 md:pt-4 md:pb-5')

    # Loop through each restaurant/event and extract relevant details
    for restaurant in restaurants:
        attendees = restaurant.find('div', class_='text-sm text-gray6')
        
        # Extract name
        name_tag = restaurant.find('h2', class_='text-gray7 font-medium text-base pb-1 pt-0 line-clamp-3')
        name = name_tag.text.strip() if name_tag else "N/A"
        
        # Extract event URL
        event_url = restaurant.find('a', class_='w-full cursor-pointer hover:no-underline')['href']

        eventResponse = requests.get(event_url)
        event_soup = BeautifulSoup(eventResponse.text, 'html.parser')
        
        # Extract date and time
        time_tag = restaurant.find('time', {'datetime': True})
        if time_tag:
            date_time_attr = time_tag['datetime']  # Get the datetime attribute
            date_time_text = time_tag.get_text(" ", strip=True)  # Extracts the readable date-time text
        else:
            date_time_text = "No date available"

        # Extract location
        location_tag = event_soup.find('div', class_='text-gray6', attrs={'data-testid': 'location-info'})
        location = location_tag.text.strip() if location_tag else "Location not available"

        # Extract hosted by
        host_name_tag = event_soup.find('div', class_='ml-6')
        host_name = "Host Name not available"
        if host_name_tag:
            host_name_tag = host_name_tag.find('span', class_='font-medium')
            host_name = host_name_tag.text.strip()

        # Extract tags
        tags = [a.get_text() for a in event_soup.find_all('a', class_='tag--topic')]

        # # Find the 'attendees' div by id
        attendees_div = event_soup.find('div', id='attendees')
        print(attendees_div)
        break


        # Extract event details
        event_details = soup.get_text(separator=" ").strip()

        # Print extracted information
        print(f"Name: {name}")
        print(f"Event URL: {event_url}")
        print(f"Date & Time: {date_time_text}")
        print(f"Location: {location}")
        print(f"Host Name: {host_name}")
        print(f"Event Tags: {tags}")
        # print(f"Number of Attendees: {attendees_count}")
        # print(event_details)

        # Export data to Excel
        # ws.append([name, event_url, date_time_text, location, host_name, ", ".join(tags)])

        print("=" * 40)

    # Save the Excel file
    # wb.save("C:/Users/DELL/OneDrive - Cambrian College/College/Capstone Project/Event_Details.xlsx")

else:
    print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
