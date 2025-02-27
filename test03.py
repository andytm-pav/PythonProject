import requests
from bs4 import BeautifulSoup


def find_rss_feed(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")
        rss_links = []

        for link in soup.find_all("link", type="application/rss+xml"):
            rss_links.append(link.get("href"))

        return rss_links if rss_links else "RSS не найден"
    except Exception as e:
        return f"Ошибка: {e}"


# Пример использования
site_url = "http://www.yandex.ru/"
print(find_rss_feed(site_url))