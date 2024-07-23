from pyrfc import Connection
import pandas as pd
from datetime import datetime, timedelta

# Connection parameters
conn_params = {
    'user': 'your_username',
    'passwd': 'your_password',
    'ashost': 'your_sap_server',
    'sysnr': '00',
    'client': '100',
    'lang': 'EN'
}

# Establish connection
conn = Connection(**conn_params)

# Define date range
today = datetime.today()
future_date = today + timedelta(days=30)  # Adjust the number of days as needed

today_str = today.strftime('%Y%m%d')
future_date_str = future_date.strftime('%Y%m%d')

# Call MB5M function module
result = conn.call('MB5M', 
                   MATNR='*', 
                   WERKS='your_plant', 
                   LGORT='your_storage_location', 
                   BUDAT= {
                       'SIGN': 'I', 
                       'OPTION': 'BT', 
                       'LOW': today_str, 
                       'HIGH': future_date_str
                   }
                  )

# Process results
if 'EXPIRY_DATE_LIST' in result:
    data = result['EXPIRY_DATE_LIST']
    df = pd.DataFrame(data)
    df.columns = ['Material', 'Batch', 'Plant', 'Storage Location', 'Expiry Date', 'Quantity']
    df['Expiry Date'] = pd.to_datetime(df['Expiry Date'])
    df.to_excel('items_about_to_expire.xlsx', index=False)
else:
    print("No data found.")
