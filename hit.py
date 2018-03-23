#!/usr/bin/env python
# encoding:utf-8

import requests
from bs4 import BeautifulSoup

SESSION = requests.session()


def sso_login(username, password):
    url = "https://ids.hit.edu.cn/authserver/login?service=http%3A%2F%2Fjwts.hit.edu.cn%2FloginCAS"

    response = SESSION.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    inputs = soup.find('form', id="casLoginForm").find_all('input', type='hidden')
    data = {}

    for i in inputs:
        data[i['name']] = i['value']

    data['username'] = username
    data['password'] = password

    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.82 Safari/537.36'
    }

    response = SESSION.post(url, headers=header, data=data)
    soup = BeautifulSoup(response.text, 'lxml')
    result = soup.find(id='msg')
    if result:
        print "[-] %s" % (result.text)
        return False
    else:
        hola = soup.find(class_='welcome')
        print "[+] %s" % (hola.text)
        return True


def download_schedule(semester_year, semester, filename):
    url = "http://jwts.hit.edu.cn/kbcx/ExportGrKbxx"
    data = {
        "xnxq": "%d-%d%d" % (semester_year, semester_year + 1, semester)
    }
    response = SESSION.post(url, data)
    with open(filename, "wb") as f:
        f.write(response.content)
    return True


def main():
    pass


if __name__ == '__main__':
    main()
