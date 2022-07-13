from chrome_driver import Driver
import os
from time import sleep
import sys
import pandas
import logging

local_path = os.path.dirname(sys.argv[0])


def main():
    hashtag_list = pandas.read_excel(os.path.join(local_path, 'Input', 'Master.xlsx'), sheet_name="Info", header=3,
                                     usecols=["Hashtag", "Status"], engine="openpyxl")
    hashtag_list.fillna("", inplace=True)
    hashtag_list = hashtag_list[(hashtag_list['Hashtag'] != "") & (hashtag_list['Status'] == "ON")]['Hashtag'].unique()
    # print(hashtag_list)
    user_list = pandas.read_excel(os.path.join(local_path, 'Input', 'Master.xlsx'), sheet_name="Info", header=3,
                                  usecols=["Name", "ID"], engine="openpyxl")
    user_list.fillna("", inplace=True)
    user_list = user_list[(user_list['Name'] != "")]
    user_list = {user['ID']: {**user, 'Hashtag': []} for user in user_list.to_dict(orient="records")}
    # print(user_list)

    os.system("TASKKILL /im chrome.exe /f /t")
    driver = Driver()
    driver.set_chrome(load_cookies=True)
    driver.get('https://www.instagram.com/')
    sleep(5)

    cookies = driver.driver_.get_cookies()
    driver.driver_.quit()
    driver = Driver()
    driver.set_chrome(load_cookies=False)
    driver.get('https://www.instagram.com/')
    sleep(5)
    # out_put = dict.fromkeys(hashtag_list, [])
    # print(out_put)

    for cookie in cookies:
        driver.driver_.add_cookie(cookie)
    logging.info('Login OK !!')
    for hashtag in hashtag_list:
        logging.info(hashtag + ' >> Start!!!')
        driver.get(f"https://www.instagram.com/explore/tags/{hashtag}/")
        xpath = "//div[text()='人気投稿']/parent::h2/following-sibling::div//a[@href] | //*[text()='このページはご利用いただけません。']"
        sleep(10)
        try:
            driver.wait_visible(xpath)
        except:
            driver.refresh()
            sleep(10)
            driver.wait_visible(xpath)
            logging.info('thu lay 人気投稿 lan 2')
        post_list = driver.get_elements("//div[text()='人気投稿']/parent::h2/following-sibling::div//a[@href]")
        if not post_list:
            logging.info("このページはご利用いただけません")
            continue
        link_list = [post.get_attribute("href") for post in post_list]
        for link in link_list:
            driver.get(link)
            xpath_1 = "//article[@role='presentation']//header//a[@href]"
            sleep(4)
            try:
                driver.wait_visible(xpath_1)
            except:
                driver.refresh()
                sleep(10)
                driver.wait_visible(xpath_1)
                logging.info('thu lay href lan 2')
            profile_name = driver.wait_visible(xpath_1).get_attribute("href")
            profile_name = str(profile_name).strip('/').split('/')[-1]
            if profile_name in user_list.keys():
                user_list[profile_name]["Hashtag"].append(hashtag)
            sleep(2)
        logging.info(hashtag + ' >> Done!!!')
        sleep(2)
    logging.info('Get Value OK !!')
    driver.driver_.quit()

    for k, v in user_list.items():
        v['Hashtag'] = chr(10).join(set(v['Hashtag'])) if len(v['Hashtag']) > 0 else "--"
    data = pandas.DataFrame(user_list.values())
    data.to_csv(os.path.join(local_path, 'Output', 'Output.csv'), encoding="UTF-16", sep="\t", index=False)
    logging.info('Write File DONE !!')
    # print(user_list)


if __name__ == '__main__':
    main()
    input("\n\nNhấn phím Enter để thoát!")
    '''os.system("TASKKILL /im cmd.exe /f /t")'''
