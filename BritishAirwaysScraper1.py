from bs4 import BeautifulSoup
import requests, openpyxl
import re

"""
     Scraping third party website https://www.airlinequality.com/
     Aim to collect customer review about British Airways
     The script can be set to be update the Xlsx files every week by importing timer 
"""

#Open empty excel sheet 
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'British Airways reviews'
#set the Review details as colum names
sheet.append(['CustomerName', 'ReviewdatePublished', 'CustomerCountry' ,'CustomerWriteUp', 'TypeofTraveller', 
              'Aircraft', 'SeatType', 'Route', 'DateFlown', 'SeatComfort', 'CabinStaffService', 
              'FoodBeverages', 'InflightEntertainment', 'GroundService', 'ValueForMoney', 'WifiConnectivity','Recommended' ])

try:
     # Set the URL of the first page
     url = 'https://www.airlinequality.com/airline-reviews/british-airways/page/'
     # Set the headers for the request
     headers = {
     'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
     # Run through pages
     for page in range(1, 350):
          response = requests.get(url + str(page))
          response.raise_for_status()
          # Parse the HTML content of the page with BeautifulSoup
          soup = BeautifulSoup(response.text, 'html.parser')
          # Find all the review boxes on the page
          review_table = soup.find('article', class_='comp comp_reviews-airline querylist position-content')
          review_boxes = review_table.find_all('article')

          ##Collect Avaliabe Data From Each Page
          for index, review_box1 in enumerate(review_boxes):
               #CustomerWriteUp = review_box1.find('div', class_= 'text_content', itemprop='reviewBody')
               customer_name_row = review_box1.find('span', itemprop='name')
               CustomerName = customer_name_row.text.strip()
               ReviewdatePublished = review_box1.find('time', itemprop="datePublished").text
               CustomerNameDetails = review_box1.find('h3', class_="text_sub_header userStatusWrapper").get_text(strip=True)
               match = re.match(r'(.+)\((.+)\)', CustomerNameDetails)
               try:
                    CustomerCountry = match.group(2)
               except:
                    CustomerCountry =None
               table = review_box1.table.text

               if 'Type Of Traveller' in table:
                    type_of_traveller_row = review_box1.table.find('td', class_='review-rating-header type_of_traveller')
                    TypeofTraveller = type_of_traveller_row.find_next_sibling('td').text.strip()
               else:
                    TypeofTraveller = None

               if 'Aircraft' in table:
                    aircraft_row = review_box1.table.find('td', class_='review-rating-header aircraft')
                    Aircraft= aircraft_row.find_next_sibling('td').text.strip()
               else:
                    Aircraft = None

               if 'Seat Type' in table:
                    seat_type_row = review_box1.table.find('td', class_='review-rating-header cabin_flown')
                    SeatType = seat_type_row.find_next_sibling('td').text.strip() 
               else: 
                    SeatType = None

               if 'Route' in table:
                    route_row = review_box1.table.find('td', class_='review-rating-header route')
                    Route = route_row.find_next_sibling('td').text.strip() 
               else: 
                    Route = None

               if 'Date Flown' in table:
                    date_flown_row = review_box1.table.find('td', class_='review-rating-header date_flown')
                    DateFlown = date_flown_row.find_next_sibling('td').text.strip() 
               else: 
                    DateFlown = None

               if 'Seat Comfort' in table:
                    seat_comfort_row = review_box1.table.find('td', class_='review-rating-header seat_comfort')
                    SeatComfort_raw = seat_comfort_row.find_next_sibling('td')
                    SeatComfort = SeatComfort_raw.find_all('span', class_="star fill")[-1].text
               else: 
                    SeatComfort = None
               if 'Cabin Staff Service' in table:
                    cabin_staff_service_row = review_box1.table.find('td', class_='review-rating-header cabin_staff_service')
                    CabinStaff_raw = cabin_staff_service_row.find_next_sibling('td')
                    CabinStaffService = CabinStaff_raw.find_all('span', class_="star fill")[-1].text
               else: 
                    CabinStaffService = None

               try:
                    food_and_beverages_row = review_box1.table.find('td', class_='review-rating-header food_and_beverages')
                    FoodBeverages_raw = food_and_beverages_row.find_next_sibling('td')
                    FoodBeverages = FoodBeverages_raw.find_all('span', class_="star fill")[-1].text
               except: 
                    FoodBeverages = None

               try:
                    inflight_entertainment_row = review_box1.table.find('td', class_='review-rating-header inflight_entertainment')
                    InflightEntertainment_raw = inflight_entertainment_row.find_next_sibling('td')
                    InflightEntertainment = InflightEntertainment_raw.find_all('span', class_="star fill")[-1].text
               except: 
                    InflightEntertainment = None

               if 'Ground Service' in table:
                    ground_service_row = review_box1.table.find('td', class_='review-rating-header ground_service')
                    GroundService_raw = ground_service_row.find_next_sibling('td')
                    GroundService = GroundService_raw.find_all('span', class_="star fill")[-1].text
               else: 
                    GroundService = None

               if 'Value For Money' in table:
                    value_for_money_row = review_box1.table.find('td', class_='review-rating-header value_for_money')
                    ValueForMoney_raw = value_for_money_row.find_next_sibling('td')
                    ValueForMoney = ValueForMoney_raw.find_all('span', class_="star fill")[-1].text
               else: 
                    ValueForMoney = None

               if 'Wifi & Connectivity' in table:
                    wifi_and_connectivity_row = review_box1.table.find('td', class_='review-rating-header wifi_and_connectivity')
                    WifiConnectivity_raw = wifi_and_connectivity_row.find_next_sibling('td')
                    WifiConnectivity = WifiConnectivity_raw.find_all('span', class_="star fill")[-1].text
               else: 
                    WifiConnectivity = None

               if 'Recommended' in table:
                    recommended_row = review_box1.table.find('td', class_='review-rating-header recommended')
                    Recommended = recommended_row.find_next_sibling('td').text.strip() 
               else: 
                    Recommended = None

               sheet.append([CustomerName, ReviewdatePublished, CustomerCountry ,TypeofTraveller, 
                         Aircraft, SeatType, Route, DateFlown, SeatComfort, CabinStaffService, 
                         FoodBeverages, InflightEntertainment, GroundService, ValueForMoney, WifiConnectivity, Recommended])
except Exception as e:
    print(e)

excel.save('British Airways reviews from third party website.xlsx')

