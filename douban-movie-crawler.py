from selenium import webdriver
from selenium.webdriver.common.by import By
from urllib.parse import unquote, urlparse, parse_qs
import requests
import json
from openpyxl import load_workbook


# 将电影数据写入excel
def write_excel_data(data):
    # 加载excel
    wb = load_workbook(filename='douban-movie-crawler.xlsx')
    # 激活excel表
    sheet = wb.active
    # 创建表头列表
    header_list = []
    # 定制表头
    for header in data[0].keys():
        if header == 'title':
            header_list.append(header)
        if header == 'actors':
            header_list.append(header)
        if header == 'release_date':
            header_list.append(header)
        if header == 'score':
            header_list.append('average')
        if header == 'types':
            header_list.append('genre')
        if header == 'url':
            header_list.append('link')
        if header == 'vote_count':
            header_list.append('votes')
    # 写入表头
    sheet.append(header_list)
    # 批量获取数据并插入到表格中
    for movie in data:
        # 用于存储最终数据
        values = []
        for header in header_list:
            if header == 'title':
                value = movie[header]
            elif header == 'actors':
                value = movie[header]
                # 列表转为字符串,并以","号分隔
                value = ','.join(value)
            elif header == 'release_date':
                value = movie[header]
            elif header == 'average':
                value = movie['score']
            elif header == 'genre':
                value = movie['types']
                # 列表转为字符串,并以","号分隔
                value = ','.join(value)
            elif header == 'link':
                value = movie['url']
            elif header == 'votes':
                value = movie['vote_count']
            else:
                value = ''
            values.append(value)
        # 过滤空字符串
        values = [x for x in values if x != '']
        # 写入excel
        sheet.append(values)
    # 保存
    wb.save('douban-movie-crawler.xlsx')
    # 关流
    wb.close()


# 获取电影数据
def get_movie_data(sort_id, limit):
    url = 'https://movie.douban.com/j/chart/top_list?type=' + sort_id + \
          '&interval_id=100:90&action=None&start=0&limit=' + limit
    try:
        # 设置请求头
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'}
        response = requests.get(url, headers=headers)
        # 检查响应状态码，如果不是 200, 则抛出异常
        response.raise_for_status()
        # 解析响应数据为 JSON 格式
        json_list = response.json()
        # 将列表转换为 JSON 字符串
        json_data = json.dumps(json_list)
        # 解析 JSON 数据为 Python 对象
        data = json.loads(json_data)
        # 返回电影数据
        return data
    except requests.exceptions.RequestException as e:
        print('获取电影数据失败:', str(e))
        return None


# 打印电影类别
def print_movie_sorts(movie_sorts):
    # 将类别按每行最多六个输出
    line_count = 0
    line = ''
    for key in movie_sorts:
        line += key + ', '
        line_count += 1
        if line_count == 6:
            print(line)
            line = ''
            line_count = 0
    if line:
        # 去除最后一个,号
        print(line.rstrip(', '))


# 获取电影类别
def get_movie_sorts():
    # 创建 Chrome WebDriver 实例
    driver = webdriver.Chrome()
    # 打开豆瓣电影排行榜页面
    driver.get("https://movie.douban.com/chart")
    # 建立类别字典
    sorts = {}
    # 获取所有电影类别的链接
    types_links = driver.find_elements(By.XPATH, '//div[@class="types"]/span/a')
    # 遍历每个类别链接
    for link in types_links:
        href = link.get_attribute("href")
        # 解码 URL
        decoded_href = unquote(href)
        # 解析 URL
        parsed_url = urlparse(decoded_href)
        # 获取查询参数
        query_params = parse_qs(parsed_url.query)
        # 获取类别 ID
        sort_id = query_params.get('type', '')[0]
        # 获取类别名称
        sort_name = query_params.get('type_name', '')[0]
        # 将类别名称和 ID 添加到字典中
        sorts.setdefault(str(sort_name), str(sort_id))
    # 关闭 WebDriver
    driver.quit()
    # 返回类别字典
    return sorts


if __name__ == '__main__':
    print("----------获取豆瓣电影排行榜数据----------")
    # 获取电影类别
    movie_sorts = get_movie_sorts()
    # 打印电影类别
    print_movie_sorts(movie_sorts)
    while True:
        # 电影类别
        sort_name = str(input('请输入电影类别：')).strip()
        if sort_name in movie_sorts.keys():
            break
        else:
            print('无效的电影类别，请重新输入！')
    # 电影总数
    limit = str(input('请输入电影总数：'))
    # 获取电影类别id
    sort_id = movie_sorts[sort_name]
    # 获取电影数据
    data = get_movie_data(sort_id, limit)
    # 将电影数据写入Excel
    write_excel_data(data)
    # 结束
    print("----------爬取成功, 程序结束----------")