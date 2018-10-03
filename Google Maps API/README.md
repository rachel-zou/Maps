# Project Purpose

In my daily work, I often need to calculate the distance and driving time from a list of zip codes to one of our e-Commerce stores.
This information can be used to determine the service area for an e-Commerce store. In this script, I utilized Google Maps API to
gather such information, including the distance and driving time in traffic during different times of the day between any zip codes 
and one of our e-Commerce stores.

# Setup

You need to configure your parameters in param.xlsx before you run the code. 
Currently, this configuration workbook contains 4 tabs.
- addr: Enter the store ids and store addresses for your stores.
- api: You should obtain an API key from Google Developer Website and enter it here (cell A2).
- time: You can specify the specific time of the day you want to use as departure time (B2:B4). Date will be next Monday.
- zip: Enter the zip codes you want to calculate in column A.

# Running the code

Run maps_api_requests.py. It will prompt you to enter the store id which will be your origin address. 
Enter the store id and hit Enter. An Excel workbook will be saved with the store id as the workbook name, 
which contains the distance from this store to each zip code in the configuration workbook,
as well as the driving time from this store to each zip code during different times of the day.
An average driving time is also provided by averaging the morning, noon and afternoon driving time.  

# Disclaimer

The script doesn't contain any confidential information. I used some random addresses in the param workbook to display the format of the workbook.
You will need your own Google API Key to run the code.  
 
