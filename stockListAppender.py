from bs4 import BeautifulSoup
import requests
import csv
import re
from datetime import datetime

base_url = "https://emiten.kontan.co.id/daftar-emiten/halaman/"
page_number = 1
stock_name = []
stock_id = []
output_file = "stocklist.csv"
debug_stocklist_file = "debug_stocklist.txt"


with open(debug_stocklist_file, mode='w') as debug_stocklist:
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    debug_stocklist.write(current_datetime,"\n")
    
while page_number <= 186:
    url_complete = f"{base_url}{page_number}"
    response = requests.get(url_complete)

    # Check if the page exists
    if response.status_code == 200:
        # Parse the HTML content
        target = BeautifulSoup(response.text, 'html.parser')

        # Extract and process the data from the current page
        raw_data = target.find_all(lambda tag: tag.name == 'a' and tag.has_attr('href') and tag.has_attr('title'))

        for element in raw_data:
            str_element = str(element)
            # Remove "(persero)" from str_element
            str_element = str_element.replace("(Persero)", "")
            if "(" in str_element:
                title_match = re.search('title="([^"]+)"', str_element)
                if title_match:
                    title_content = title_match.group(1)
                    
                    title_parts = title_content.split(" (")
                    print("element found", title_parts)

                    # Print title_parts on a new line in debug.txt
                    with open(debug_stocklist_file, mode='a') as txt_file:
                        txt_file.write(" | ".join(map(str,title_parts)) + "\n")
                    
                    data_1_temp = title_parts[0]
                    data_2_temp = title_parts[1][:-1]  # Exclude the closing parenthesis
                    # Check if the last data_1 array is the same as data_1_temp
                    if not stock_name or stock_name[-1] != data_1_temp:
                        stock_name.append(data_1_temp)
                        stock_id.append(data_2_temp)
                else:
                    print("element not found in page: ", page_number)

        # Move on to the next page
        page_number += 1
        print("Done page: ", page_number)
    else:
        print("done/error 404")
        break

# test output
print(stock_id)

# combine the stock name and stock id
combined_data = zip(stock_id, stock_name)

# print to csv file
with open(output_file, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(combined_data)
