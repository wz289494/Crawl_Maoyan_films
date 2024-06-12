import requests
import re
import json
import os
import pandas as pd

def set_cookies():
    """
    Set cookies for the request.

    Returns:
    dict: A dictionary of cookies.
    """
    cookies = {
        'Hm_lvt_703e94591e87be68cc8da0da7cbd0be2': '1712716289',
        '_lxsdk_cuid': '18ec5d8956dc8-0458d307378267-26001a51-144000-18ec5d8956dc8',
        '_lxsdk': '6BC0BFE0F6E211EE9F28D3BFFA06DA49A4F7C0C2B1774E3BA76F01CC7AB90F48',
        'Hm_lpvt_703e94591e87be68cc8da0da7cbd0be2': '1712716309',
        '_lx_utm': 'utm_source%3Dbing%26utm_medium%3Dorganic',
        '_lxsdk_s': '18ec5d8956e-94e-2fc-ffa%7C%7C9',
    }
    return cookies

def set_headers():
    """
    Set headers for the request.

    Returns:
    dict: A dictionary of headers.
    """
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Referer': 'https://piaofang.maoyan.com/rankings/year',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'Uid': 'd29e510e35171d84b88710fc38f91c8ff119b947',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }
    return headers

def set_params(year, tab):
    """
    Set parameters for the request.

    Parameters:
    year (int): The year to query.
    tab (int): The tab number to query.

    Returns:
    dict: A dictionary of parameters.
    """
    params = {
        'year': year,
        'limit': '100',
        'tab': tab,
    }
    return params

def get_pages(cookies, headers, params):
    """
    Get the page content from the website.

    Parameters:
    cookies (dict): Cookies for the request.
    headers (dict): Headers for the request.
    params (dict): Parameters for the request.

    Returns:
    str: The page content.
    """
    response = requests.get('https://piaofang.maoyan.com/rankings/year', params=params, cookies=cookies, headers=headers)
    return response.text

def hander_page(data):
    """
    Handle the page content and extract movie data.

    Parameters:
    data (dict): The JSON data from the page.

    Returns:
    list: A list of dictionaries containing movie data.
    """
    html = data['yearList']
    pattern = r'hrefTo,href:(.*?)>.*?class="first-line">(.*?)</p>.*?second-line">(.*?) 上映</p>.*?col2 tr">(.*?)</li>.*?col3 tr">(.*?)</li>.*?col4 tr">(.*?)</li>'
    matches = re.findall(pattern, html, re.DOTALL)

    movie_data = []
    for match in matches:
        movie_data.append({
            '电影链接': match[0],
            '电影名称': match[1],
            '上映日期': match[2],
            '电影票房': int(match[3]),
            '平均票价': float(match[4]),
            '场均上座率': int(match[5])
        })
    return movie_data

def create_dataframe(data_info):
    """
    Create a pandas DataFrame from the movie data.

    Parameters:
    data_info (list): A list of dictionaries containing movie data.

    Returns:
    pandas.DataFrame: The created DataFrame.
    """
    df = pd.DataFrame(data_info)
    return df

def df_to_excel(df, output_path):
    """
    Save the DataFrame to an Excel file.

    Parameters:
    df (pandas.DataFrame): The DataFrame to save.
    output_path (str): The path to the output Excel file.

    Returns:
    None
    """
    # If the file does not exist, create a new one
    if not os.path.exists(output_path):
        df.to_excel(output_path, index=False)
    else:
        with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            # Get the maximum row number of the existing sheet
            start_row = writer.sheets['Sheet1'].max_row
            # If there is existing content (i.e., max row > 1), avoid writing the header again
            if start_row > 1:
                header = False
            else:
                header = True
            # Write the DataFrame, avoiding additional headers
            df.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False, header=header)

def main():
    """
    Main function to execute the scraping and data processing.

    Returns:
    None
    """
    # Set request headers and cookies
    cookies = set_cookies()
    headers = set_headers()
    print('Parameters set successfully')

    # Loop through each year and scrape data
    for year in range(2011, 2024):
        number = 2024 - (year - 1)

        # Set parameters for the request
        params = set_params(year, number)

        # Get the page content
        pages = get_pages(cookies, headers, params)
        print(f'Currently scraping year {year}')

        data = json.loads(pages)

        # Handle and parse the page content
        new_data = hander_page(data)
        print(new_data)
        print('Parsing completed')

        # Create a DataFrame and save to Excel
        df = create_dataframe(new_data)
        df_to_excel(df, '猫眼电影_11to23.xlsx')
        print('Data added to Excel')

if __name__ == '__main__':
    main()
