import asyncio
import aiohttp
from bs4 import BeautifulSoup
import pandas as pd


async def fetch_page(session, page_number):
    base_url = 'https://www.gulfoodgreen.com/exhibitors'
    params = {
        'page': page_number,
        'searchgroup': '00000001-exhibitors'
    }
    headers = {
        'User-Agent': 'Mozilla/5.0'
    }
    async with session.get(base_url, params=params, headers=headers) as response:
        if response.status == 200:
            text = await response.text()
            return text
        else:
            print(f"Failed to fetch page {page_number}")
            return None

def parse_page(html):
    soup = BeautifulSoup(html, 'html.parser')
    data = []

    exhibitors_list = soup.find('ul', class_='m-exhibitors-list__items js-library-list')
    if not exhibitors_list:
        print("No exhibitors list found.")
        return data

    exhibitors = exhibitors_list.find_all('li', class_='m-exhibitors-list__items__item')

    for exhibitor in exhibitors:
        # Exhibitor Name
        name_tag = exhibitor.find('h2', class_='m-exhibitors-list__items__item__name')
        name = name_tag.get_text(strip=True) if name_tag else ''

        # Sectors
        sectors_tag = exhibitor.find('div', class_='m-exhibitors-list__items__item__sectors')
        sectors = sectors_tag.get_text(strip=True) if sectors_tag else ''

        # Country
        country_tag = exhibitor.find('div', class_='m-exhibitors-list__items__item__location')
        country = country_tag.get_text(strip=True) if country_tag else ''

        data.append({
            'Exhibitor Name': name,
            'Sectors': sectors,
            'Country': country
        })

    return data

async def main():
    async with aiohttp.ClientSession() as session:
        tasks = []
        for page_number in range(1, 6):
            task = asyncio.ensure_future(fetch_page(session, page_number))
            tasks.append(task)

        pages = await asyncio.gather(*tasks)

        all_data = []
        for html in pages:
            if html:
                data = parse_page(html)
                all_data.extend(data)

        # Save data to Excel
        df = pd.DataFrame(all_data)
        df.to_excel('exhibitors.xlsx', index=False)
        print("Data saved to exhibitors.xlsx")

if __name__ == '__main__':
    asyncio.run(main())
