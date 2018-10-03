import requests
import pandas as pd
import sys
import maps_api_config

# Prompt the user to enter a store id
store_id = int(input("Enter store id: "))

# Check if the user has entered a valid store id
try:
    store_addr = maps_api_config.store_address(store_id)
except KeyError:
    print ("Please enter a valid store id.")
    sys.exit(1)

# Define the function to ping Google Maps API
def google_maps_api(shift):
    '''
    Input: shift name - morning, noon or afternoon. 
    Output: data returned from Google Maps API.
    '''
    url = maps_api_config.url(store_id, shift)
    response = requests.get(url)
    data = response.json()
    if data["status"] == 'OK':
        return data
    else:
        print ("Please make sure you have a valid Google API Key.")
        sys.exit(1)

# Ping Google API and return the traffic information in the morning
data = google_maps_api('morning')
size = len(data["rows"][0]['elements'])

# Create the empty lists to store your result        
From_list = []
From_list.append(store_id)
valid_zip = []
invalid_zip = []
distance_value = []
duration_value = []
duration_in_trafic_value = []

for i in range(size):
    if data["rows"][0]['elements'][i]['status'] == 'OK':
        distance_value.append(float("{0:.1f}".format(data["rows"][0]['elements'][i]['distance']['value']*0.00062137)))
        duration_value.append(float("{0:.1f}".format(data["rows"][0]['elements'][i]['duration']['value']/60)))
        duration_in_trafic_value.append(float("{0:.1f}".format(data["rows"][0]['elements'][i]['duration_in_traffic']['value']/60)))
        valid_zip.append(maps_api_config.ziplist()[i])
    else:
        invalid_zip.append(maps_api_config.ziplist()[i])
        continue
        
zip_value = {'From': From_list * len(valid_zip),
             'To': valid_zip,
             'Distance (miles)': distance_value,
             'Duration (mins)': duration_value,
             'Morning (mins)': duration_in_trafic_value}
        
df_morning = pd.DataFrame.from_dict(zip_value)

# Ping Google API and return the traffic information around noon 
data = google_maps_api('noon')
size = len(data["rows"][0]['elements'])

duration_in_trafic_value = []

for i in range(size):
    if data["rows"][0]['elements'][i]['status'] == 'OK':        
        duration_in_trafic_value.append(float("{0:.1f}".format(data["rows"][0]['elements'][i]['duration_in_traffic']['value']/60)))
    else:
        continue
        
zip_value = {'Noon (mins)': duration_in_trafic_value}
df_noon = pd.DataFrame.from_dict(zip_value)['Noon (mins)']

# Ping Google API and return the traffic information in the afternoon
data = google_maps_api('afternoon')
size = len(data["rows"][0]['elements'])

duration_in_trafic_value = []

for i in range(size):
    if data["rows"][0]['elements'][i]['status'] == 'OK':        
        duration_in_trafic_value.append(float("{0:.1f}".format(data["rows"][0]['elements'][i]['duration_in_traffic']['value']/60)))
    else:
        continue
        
zip_value = {'Afternoon (mins)': duration_in_trafic_value}
df_afternoon = pd.DataFrame.from_dict(zip_value)['Afternoon (mins)']

# Combine the traffic data during different times of the day into one data frame
result = pd.concat([df_morning,df_noon,df_afternoon], axis=1)
result['Average (mins)'] = (result['Morning (mins)']+result['Noon (mins)']+result['Afternoon (mins)'])/3
result['Average (mins)'] = result['Average (mins)'].round(1)

# Store any invalid zip codes in a separate data frame
invalid = [('invalid zip', invalid_zip)]
df_invalid = pd.DataFrame.from_items(invalid) 

# Write the result in a spreadsheet and save it
writer = pd.ExcelWriter('%d.xlsx' % store_id)
result.to_excel(writer, '%d' % store_id, index=False, columns=['From', 'To', 'Distance (miles)', 'Duration (mins)', 'Morning (mins)', 'Noon (mins)', 'Afternoon (mins)', 'Average (mins)'])
df_invalid.to_excel(writer, 'invalid_zip', index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['%d' % store_id]

# Set the column width and format.
worksheet.set_column('C:H', 18)

writer.save()