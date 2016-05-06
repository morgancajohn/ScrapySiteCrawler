import os
import glob
import time
import datetime
from contextlib import contextmanager, closing

import MySQLdb
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By


MYSQL_HOST = 'localhost'
MYSQL_DBNAME = 'fedex'
MYSQL_USER = 'fedex'
MYSQL_PASSWD = ''
# MYSQL_USER = 'root'
# MYSQL_PASSWD = 'rootroot'

db = MySQLdb.connect(
    host=MYSQL_HOST,
    db=MYSQL_DBNAME,
    user=MYSQL_USER,
    passwd=MYSQL_PASSWD,
    charset='utf8',
    use_unicode=True,
)


def main():
    for claim in fetch_claims():
        # Process
        with Driver() as driver:
            try:
                completed, note = scenario(driver, claim)
            except:
                completed, note = (False, "Failed")


def fetch_claims():
    # Fetch and lock claims
    with db as cursor:
        claims = fetch_dicts(cursor, '''
            select * from ups_account
            where 1
        ''')
        return claims


def scenario(driver, claim):
    note = 'Failed'
    completed = False

    driver.delete_all_cookies()
    driver.get('https://www.apps.ups.com/ebilling/invoice/companysummarySearch.do?reportId=upsCompanySumm&status=initial')

    # Log in
    username = driver.find_element_by_css_selector('[name=uid]')
    username.send_keys(claim['username'])
    password = driver.find_element_by_css_selector('[name=password]')
    password.send_keys(claim['password'])
    login_btn = driver.find_element_by_css_selector('[id=submitBtn]')
    login_btn.click()
    time.sleep(5)

    # Show search form
    driver.get('https://www.apps.ups.com/ebilling/invoice/companysummarySearch.do?reportId=upsCompanySumm&status=initial')
    # today = datetime.datetime.today()
    today = datetime.datetime(2015, 5, 10, 5, 42, 50, 324950)
    datefrom = today - datetime.timedelta(days = 7)
    str_datefrom = datefrom.strftime('%m/%d/%Y')
    str_today = today.strftime('%m/%d/%Y')

    fromStatementDate = driver.find_element_by_css_selector('[name=fromStatementDate]')
    fromStatementDate.send_keys(str_datefrom)
    toStatementDate = driver.find_element_by_css_selector('[name=toStatementDate]')
    toStatementDate.send_keys(str_today)

    select = Select(driver.find_element_by_css_selector('[name=invStatus]'))
    select.select_by_visible_text('All')

    select = Select(driver.find_element_by_css_selector('[name=acctNumber]'))
    select.select_by_visible_text('All')

    try:
        search = driver.find_elements(By.XPATH, '//input[@title="Search"]')[0]
        search.click()
    except:
        return completed, note
    time.sleep(5)

    a_tags = driver.find_elements(By.XPATH, '//div[@class="mc2"]//table[@class="border"]/tbody/tr/td[3]/a')
    for a_tag in a_tags:
        href = a_tag.get_attribute('href')
        driver.get(href)
        time.sleep(5)
        driver.find_element_by_id('viewAndDownloadInvoiceBtn').click()
        time.sleep(2)

    try:
        radios = driver.find_elements(By.XPATH, '//input[@type="radio"]')
        for radio in radios:
            if radio.get_attribute("value") == "csv":
                radio.click()
    except:
        return completed, note

    driver.find_element_by_id('viewAndDownloadInvoiceDialogLiteCsvSubmitBtn').click()
    time.sleep(10)
    # <username>-<date>.csv
    newest = max(glob.glob('%s/data/*.*' % os.getcwd()), key=os.path.getctime)
    dst = "%s/data/%s-%s.csv" % (os.getcwd(), claim['username'], today.strftime('%Y_%m_%d'))
    os.rename(newest, dst)


    note = 'Successfully done'
    # Analyze result
    completed = True
    return completed, note


@contextmanager
def Driver():
    fp = webdriver.FirefoxProfile()

    fp.set_preference("browser.download.folderList",2)
    fp.set_preference("browser.download.manager.showWhenStarting",False)
    fp.set_preference("browser.download.dir", "%s/data" % os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/csv,text/plan,text/comma-separated-values")
    fp.set_preference("browser.helperApps.alwaysAsk.force", False)
    fp.set_preference("browser.download.manager.useWindow", False)
    fp.set_preference("browser.download.manager.closeWhenDone", True)

    driver = webdriver.Firefox(firefox_profile=fp)

    yield driver
    driver.quit()


# Database utils

from operator import itemgetter

def fetch_dicts(cursor, sql, params=()):
    cursor.execute(sql, params)
    field_names = _field_names(cursor)
    return [dict(zip(field_names, row)) for row in cursor.fetchall()]

def _field_names(cursor):
    return map(itemgetter(0), cursor.description)


if __name__ == '__main__':
    main()
