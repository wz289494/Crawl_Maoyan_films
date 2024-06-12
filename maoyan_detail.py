import time
from bs4 import BeautifulSoup
import requests
import pandas as pd
import os

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
        'theme': 'moviepro',
        '_lxsdk_s': '18ec5d8956e-94e-2fc-ffa%7C%7C12',
    }
    return cookies

def set_headers():
    """
    Set headers for the request.

    Returns:
    dict: A dictionary of headers.
    """
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }
    return headers

def get_pages(cookies, headers, id):
    """
    Get the page content from the Maoyan website for a specific movie ID.

    Parameters:
    cookies (dict): Cookies for the request.
    headers (dict): Headers for the request.
    id (str): The movie ID.

    Returns:
    str: The page content.
    """
    response = requests.get(f'https://piaofang.maoyan.com/movie/{id}', cookies=cookies, headers=headers)
    return response.text

def extract_info(html):
    """
    Extract movie information from the HTML content.

    Parameters:
    html (str): The HTML content of the page.

    Returns:
    dict: A dictionary containing extracted movie information.
    """
    soup = BeautifulSoup(html, 'html.parser')
    movie_info = {}

    # Extract movie titles
    title_cn = soup.find('span', class_='info-title-content')
    title_en = soup.find('span', class_='info-etitle-content')
    movie_info['title_cn'] = title_cn.text.strip() if title_cn else None
    movie_info['title_en'] = title_en.text.strip() if title_en else None

    # Extract movie category and format
    category = soup.find('p', class_='info-category')
    movie_info['category'] = category.text.strip() if category else None

    # Extract origin and duration
    source_duration = soup.find('p', class_="ellipsis-1")
    movie_info['source_duration'] = source_duration.text.strip() if source_duration else None

    # Extract release date
    release = soup.find('span', class_='score-info')
    movie_info['release'] = release.text.strip() if release else None

    # Extract rating information
    rating_info = {}
    percentbars = soup.find_all('div', class_='percentbar')
    rating_distribution = {}
    for bar in percentbars:
        label = bar.find('span', class_='percentbar-label').text.strip()
        value = bar.find('span', class_='percentbar-val').text.strip()
        rating_distribution[label] = value
    rating_info['rating_distribution'] = rating_distribution

    rating_num = soup.find('span', class_='rating-num')
    score_count = soup.find('p', class_='detail-score-count')
    wish_count = soup.find('p', class_='detail-wish-count')
    rating_info['average_rating'] = rating_num.text.strip() if rating_num else None
    rating_info['score_count'] = score_count.text.strip() if score_count else None
    rating_info['wish_count'] = wish_count.text.strip() if wish_count else None

    movie_info['rating_info'] = rating_info

    # Extract user persona information
    persona_info = {}
    gender_info = {}
    gender_items = soup.find_all('div', class_='persona-line-item')
    for item in gender_items:
        key = item.find('div', class_='persona-item-key').text.strip()
        value = item.find('div', class_='persona-item-value').text.strip()
        gender_info[key] = value
    persona_info['gender'] = gender_info

    city_info = {}
    city_items = soup.find_all('div', class_='persona-item')
    for item in city_items:
        key = item.find('div', class_='persona-item-key').text.strip()
        value = item.find('div', class_='persona-item-value').text.strip()
        city_info[key] = value
    persona_info['city'] = city_info

    movie_info['persona_info'] = persona_info

    return movie_info

def extract_info_add(html):
    """
    Extract additional movie information such as region and duration.

    Parameters:
    html (str): The HTML content of the page.

    Returns:
    dict: A dictionary containing the region and duration of the movie.
    """
    soup = BeautifulSoup(html, 'lxml')
    info = soup.find('div', class_='info-source-duration')
    if info:
        text = info.get_text(strip=True)
        parts = text.split('/')
        if len(parts) == 2:
            region = parts[0].strip()
            duration = parts[1].strip()
            return {'region': region, 'duration': duration}
    return None

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

if __name__ == '__main__':
    # Read the Excel file containing movie links
    df = pd.read_excel('猫眼电影_11to23.xlsx')
    # Create a list of movie IDs
    movie_id_list = df['电影链接'].tolist()

    # Set cookies and headers for the requests
    cookies = set_cookies()
    headers = set_headers()

    for index, movie_id in enumerate(movie_id_list):
        # Get the page content for each movie
        pages = get_pages(cookies, headers, movie_id)

        # Extract additional movie information
        data = extract_info_add(pages)
        print(data)

        # Create a DataFrame and save to Excel
        df = pd.DataFrame([data])
        df_to_excel(df, '3.xlsx')

        # Print progress
        print(f"Progress: {index + 1}/{len(movie_id_list)} ({(index + 1) / len(movie_id_list) * 100:.2f}%)")

        time.sleep(0.5)
