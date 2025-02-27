import requests
from bs4 import BeautifulSoup
# from lxml import html

# URL главной страницы mail.ru
# url = 'https://mail.ru'
url = 'https://news.mail.ru/society/64549964/?frommail=1&utm_partner_id=625'
# url = 'https://iz.ru'
# url = 'https://www.finam.ru/'

# Получаем HTML-код страницы
response = requests.get(url)

# print(response.text)
# print(dir(response))
# print(response.text)
# print(response.content)
# tree = html.fromstring(response.content)
# print(tree)
# rss_string = tree.xpath('//rss')[0].text
#
# print(rss_string)

# Проверяем, что запрос успешен
if response.status_code == 200:
    # Парсим HTML-код
    soup = BeautifulSoup(response.text, 'html.parser')

    # Пример: извлечение заголовка страницы
    title = soup.title.string
    print(f'Заголовок страницы: {title}')

    # Здесь можно добавить дополнительный код для извлечения других данных

    # Ищем все ссылки на RSS
    rss_links = soup.find_all('link', type='application/rss+xml')

    # Выводим найденные ссылки
    for link in rss_links:
        print(link.get('href'))

else:
    print(f'Ошибка при получении страницы: {response.status_code}')
