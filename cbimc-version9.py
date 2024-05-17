# -*- coding: utf-8 -*-
"""
Created on Sun Dec 10 18:35:42 2023

@author: fdd
"""
import requests
import json
import pandas as pd
from datetime import datetime

def search_all_articles(keyword, page_size=20, channel=1, sort='date desc', site_id=1):
    url = 'http://www.cbimc.cn/xy/Search.do'
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'User-Agent': ' dongdong?'
    }

    articles = []
    page_no = 1

    while True:
        data = {
            'q': keyword,
            'pageNo': page_no,
            'pageSize': page_size,
            'channel': channel,
            'sort': sort,
            'siteID': site_id
        }

        response = requests.post(url, data=data, headers=headers)

        if response.status_code == 200:
            response_json = json.loads(response.text)
            articles_in_page = len(response_json['article'])

            if articles_in_page == 0:
                break

            articles.extend(response_json['article'])
            page_no += 1
        else:
            print(f"Error on page {page_no}: Received response code {response.status_code}")
            break

    return articles

def count_articles_after_date(articles, target_date_str):
    target_date = datetime.strptime(target_date_str, '%Y-%m-%d')
    return sum(datetime.strptime(article['date'], '%Y-%m-%d') >= target_date for article in articles)


def save_articles_to_excel(articles, output_file, target_date_str):
    df = pd.DataFrame(articles)
    df = df[['title', 'date', 'url', 'enpcontent', 'keyword']]  # Select only specific columns
    count_after_date = count_articles_after_date(articles, target_date_str)
    
    # Inserting the count in the first row
    first_row = pd.DataFrame({'title': [f'Special note: Number of articles after {target_date_str} is {count_after_date}']})
    df = pd.concat([first_row, df], ignore_index=True)
    
    # Inserting Serial Number column
    df.insert(1, 'Serial Number', range(1, 1 + len(df)))

    df.to_excel(output_file, index=False)

# 使用爬虫并保存结果到 Excel 文件
keyword = '人工智能'
output_file = 'forxiaoluo.xlsx'  # Excel 文件的输出路径
target_date_str = '2023-12-01'  # 设置目标日期
articles = search_all_articles(keyword)
save_articles_to_excel(articles, output_file, target_date_str)

print(f"Articles saved to Excel file: {output_file}")
