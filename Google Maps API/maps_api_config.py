import pandas as pd

def store_list():
    '''
    Input: None. 
    Output: A list of store IDs.
    '''
    stores = pd.read_excel('param.xlsx', sheetname = 'addr')
    store_list = stores['store'].tolist()
    return store_list

def store_address(store_id):
    '''
    Input: Store_id. 
    Output: Store address.
    '''
    addr = pd.read_excel('param.xlsx', sheetname = 'addr')
    addr = addr.set_index('store').T.to_dict('list')
    return addr[store_id][0]

def api():
    '''
    Input: None.
    Output: Your Google API key.
    '''
    api = pd.read_excel('param.xlsx', sheetname = 'api')
    return api['api'][0]

def time(shift):
    '''
    Input: Morning, noon or afternoon. 
           (By default, morning is 7 AM, noon is 11 AM and afternoon is 5 PM. Date is next Monday.
           You can change your default setting in the param.xlsx.)
    Output: Your specified departure time (shown as seconds from 1/1/1970 12 AM UTC).
    '''
    time = pd.read_excel('param.xlsx', sheetname = 'time')
    time = time.set_index('shift').T.to_dict('list')
    return int(round(time[shift][4]))

def ziplist():
    '''
    Input: None.
    Output: Zip codes in a list.
    '''
    zipcode = pd.read_excel('param.xlsx', sheetname = 'zip')
    zipcode = zipcode['zip'].tolist()
    return zipcode

def zipcode():
    '''
    Input: None.
    Output: The combined zip code you can directly pass to Google Maps API url.
    '''
    zipcode = pd.read_excel('param.xlsx', sheetname = 'zip')
    zipcode = zipcode['zip'].tolist()
    return '|'.join(map(str, zipcode))

def url(store_id, shift):
    '''
    Input: store_id, shift: 'morning', 'noon', 'afternoon'.
    Output: Google Maps API url.
    '''
    origins = store_address(store_id)
    destinations = zipcode()
    departure_time = time(shift)
    key = api()
    return 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=%s&destinations=%s&departure_time=%s&traffic_model=best_guess&key=%s' % (origins, destinations, departure_time, key)
