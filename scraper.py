

from selenium import webdriver
import urllib
from urlparse import urljoin
import selenium.webdriver as webdriver
#import urllib.request
from bs4 import BeautifulSoup
import contextlib
import lxml.html
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.by import By

driver = webdriver.PhantomJS()
driver.set_window_size(1120, 550)

driver.delete_all_cookies()

## GOING TO LINK
# driver.get("http://uwo.summon.serialssolutions.com/2.0.0/link/dbrecommender/AAAzMi4wLVNVTU1PTi1TRVNTSU9OLWNiNjNlMmFlYTkzMzIzYWU4ZDhlMzkyZDBjNmFmZmNiAQAKVlIyUEs5U1g5VwIATGh0dHBzOi8vd3d3LmxpYi51d28uY2EvY2dpLWJpbi9lenBhdXRobi5jZ2k_dXJsPWh0dHA6Ly93d3cuTWVyZ2VudE9ubGluZS5jb20DAAEwBAAOTWVyZ2VudCBPbmxpbmUFAAhkYXRhYmFzZQ")
driver.get("https://www.lib.uwo.ca/cgi-bin/ezpauthenticate.cgi?url=http://www.MergentOnline.com")
# driver.get("http://www.mergentonline.com.proxy1.lib.uwo.ca/basicsearch.php")


## LOGIN
element_user = driver.find_element_by_name("user")
element_user.send_keys("zhemani")

with file("pass.txt") as f:
    password = f.read()

element_password = driver.find_element_by_name("pass")
element_password.send_keys(password)




# WORKS
def find_by_xpath(locator):
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, locator))
    )

    return element

find_by_xpath('//input[@value = "Login"]').click()

element_ticker = driver.find_element_by_name("searchtext")

# TODO - GET FROM REQUEST
element_ticker.send_keys("msft")
driver.find_element_by_id("basicsearchbtn").click()
driver.get_screenshot_as_file('mhacks3.png')

ret = driver.find_element_by_name("chkopt[]")
parent = ret.find_element_by_xpath('..')

# Navigation
ret1 = parent.find_element_by_xpath('following-sibling::td/a')
#driver.find_element_by_xpath("//input[@name, 'chkopt[]']/following-sibling")
link1 = driver.find_element_by_link_text(ret1.text)
link1.click()
driver.get_screenshot_as_file('mhacks5.png')

link2 = driver.find_element_by_link_text("Company Financials")
link2.click()
driver.get_screenshot_as_file('mhacks6.png')

element5 = driver.find_element_by_name("range")
element5.send_keys("7")
element5.send_keys("enter")
element6 = driver.find_element_by_name("Submit").click()

driver.get_screenshot_as_file('mhacks7.png')




################################################################################
html_doc = driver.page_source

soup = BeautifulSoup(html_doc)
soup.prettify()
table = soup.find("table",{"id": "table01"})


def makelist(table):
    result = []
    allrows = table.findAll('tr')
    allrows = filter(None, allrows)
    result = filter(None, result)
    for row in allrows:
        result.append([])
        allcols = row.findAll('td')
        for col in allcols:
            thestrings = [unicode(s) for s in col.findAll(text=True)]
            thetext = ''.join(thestrings)
            result[-1].append(thetext)
            result = filter(None, result)

    data = []

    for i in range(len(result)):
        data.append([])

        for j in range (len(result[1])):
            if result[i][j] != "":
                data[i].append(result[i][j])


    return data

BS = makelist(table)




target = open("balanceSheet.csv", 'w')

for i in range (len(BS)):
    for j in range(len(BS[i])):
        target.write('"' + BS[i][j] + '"')
        target.write(",")

    target.write("\n")


############################################################

element11 = driver.find_element_by_name("dataarea")
element11.send_keys("i")
element11.send_keys("enter")
element12 = driver.find_element_by_name("Submit").click()

driver.get_screenshot_as_file('mhacks15.png')

html_doc2 = driver.page_source

soup2 = BeautifulSoup(html_doc2)
soup2.prettify()

table2 = soup2.find("table",{"id": "table02"})

IS = makelist(table2)

target = open("incomeStatement.csv", 'w')

for i in range (len(IS)):
    for j in range(len(IS[i])):
        target.write('"' + IS[i][j] + '"');
        target.write(",")

    target.write("\n")


link = driver.find_element_by_link_text("Log Out")
link.click()
driver.get_screenshot_as_file('mhacks4.png')


driver.quit()



###### REST API ######


app = Flask(__name__)
api = Api(app)

class HelloWorld(Resource):
    def get(self):
        return {'hello': 'world'}

api.add_resource(HelloWorld, '/')

if __name__ == '__main__':
    app.run(debug=True)
