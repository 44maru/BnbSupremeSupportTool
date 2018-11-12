import copy
import datetime
import json
import logging.config
import os
import sys
from datetime import datetime as dt

import requests
from xlrd import open_workbook

import licensemanager

LOG_CONF = "./logging.conf"
logging.config.fileConfig(LOG_CONF)

from bs4 import BeautifulSoup
from kivy.app import App
from kivy.clock import Clock
from kivy.config import Config

Config.set('modules', 'inspector', '')  # Inspectorを有効にする
Config.set('graphics', 'width', 1280)  # Windowの幅を1280にする
Config.set('graphics', 'maxfps', 20)  # フレームレートを最大で20にする
Config.set('graphics', 'resizable', 0)  # Windowの大きさを変えられなくする
Config.set('input', 'mouse', 'mouse,disable_multitouch')
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.core.window import Window
from kivy.resources import resource_add_path
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.screenmanager import Screen
from kivy.uix.textinput import TextInput

if hasattr(sys, "_MEIPASS"):
    resource_add_path(sys._MEIPASS)

EMPTY = ""
SIZE_S = "s"
SIZE_M = "m"
SIZE_L = "l"
SIZE_XL = "xl"
KEY_INDEX = "index"
KEY_ITEM_CODE = "itemCode"
KEY_NAME = "Name"
KEY_TITLE = "title"
KEY_MONITOR = "monitor"
KEY_SCHEDULE = "schedule"
KEY_COLOR = "color"
KEY_PAGE = "page"
KEY_CHECKOUT_DELAY_ENABLED = "checkout_delay_enabled"
KEY_CHECKOUT_DELAY_SECONDS = "checkout_delay_seconds"
KEY_PROXY_RATIO = "proxyratio"
KEY_TASK_RATIO = "taskratio"
KEY_TASKS = "tasks"
KEY_EMAIL = "email"
KEY_BILLING = "billing"
KEY_FIRST_NAME = "FirstName"
KEY_LAST_NAME = "LastName"
KEY_ADDRESS1 = "Address1"
KEY_ADDRESS2 = "Address2"
KEY_ZIP_CODE = "ZipCode"
KEY_PHONE = "Phone"
KEY_CITY = "City"
KEY_STATE = "State"
KEY_PROVINCE = "Province"
KEY_COUNTRY = "Country"
KEY_EU_REGION = "EURegion"
KEY_DSM_REGION = "DSMRegion"
BILLING_AND_SHIPPING = ["Billing", "Shipping"]

KEY_CARD_TYPE = "CreditCardType"
KEY_CARD_NUMBER = "CreditCardNumber"
KEY_CARD_EXPIRY_MONTH = "CreditCardExpiryMonth"
KEY_CARD_EXPIRY_YEAR = "CreditCardExpiryYear"
KEY_CARD_CVV = "CreditCardCvv"
KEY_MAX_CHECKOUTS = "MaxCheckouts"

ACCOUNT_KEY_ADIDAS_EXPLOIT = "AdidasExploit"
ACCOUNT_KEY_EMAIL = "EmailAddress"
ACCOUNT_KEY_PASSWORD = "Password"
ACCOUNT_KEY_SIZE = "Size"
ACCOUNT_KEY_LINKS = "Links"
ACCOUNT_KEY_KEYWORDS = "Keywords"
ACCOUNT_KEY_NOTIFICATION_EMAIL = "NotificationEmail"
ACCOUNT_KEY_NOTIFICATION_TEXT = "NotificationText"
ACCOUNT_KEY_SITE_TYPE = "SiteType"
ACCOUNT_KEY_IS_GUEST = "IsGuest"
ACCOUNT_KEY_DISABLED = "Disabled"
ACCOUNT_KEY_CHECKOUT_INFO = "CheckoutInfo"
ACCOUNT_KEY_PAY_PAL_CHECKOUT = "PayPalCheckout"
ACCOUNT_KEY_CC_CHECKOUT = "CcCheckout"
ACCOUNT_KEY_FINALIZE_ORDER = "FinalizeOrder"
ACCOUNT_KEY_PAY_PAL_EMAIL_ADDRESS = "PayPalEmailAddress"
ACCOUNT_KEY_PAY_PAL_PASSWORD = "PayPalPassword"
ACCOUNT_KEY_CC_PROFILE = "CcProfile"
ACCOUNT_KEY_ONLY_NEW_SUPREAME_PRODUCTS = "OnlyNewSupremeProducts"
ACCOUNT_KEY_CHECKOUT_DELAY = "CheckoutDelay"
ACCOUNT_KEY_QUANTITY = "Quantity"

KEY_SIZES = "sizes"
KEY_CHECKOUT_PROFILE = "checkoutprofile"
INDEX_PROFILE_NAME = 0
INDEX_FIRST_NAME = 3
INDEX_LAST_NAME = 4
INDEX_ZIP_CODE = 5
INDEX_STATE = 6
INDEX_CITY = 7
INDEX_ADDRESS = 8
INDEX_PHONE = 9
INDEX_EMAIL = 10
INDEX_PAY_TYPE = 11
INDEX_CARD_NUMBER = 12
INDEX_CARD_EXPIRY_MONTH = 13
INDEX_CARD_EXPIRY_YEAR = 14
INDEX_CARD_CVV = 15
INDEX_ITEM_NO = 1
INDEX_ITEM_SIZE = 2
VAL_CHECKOUT_DELAY_ENABLED = False
VAL_CHECKOUT_DELAY_SECONDS = 5.0
VAL_PROXY_RATIO = 1
VAL_TASK_RATIO = 1

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0",
}

ID_GET_SITE_INFO_BUTTON = "get_site_info_button"
ID_DUMP_CHECKOUTPROFILES_BUTTON = "dump_checkoutprofiles_button"
ID_DUMP_RELEASEPROFILES_BUTTON = "dump_releaseprofiles_button"
ID_MESSAGE = "message"
ID_MAX_DATA_NUM_PER_FILE = "max_data_num_per_file"
ID_BUTTON_TYPE_PC = "type_pc"
ID_BUTTON_TYPE_MOBILE = "type_mobile"

RELEASE_PROFILES_JSON = "releaseprofiles{}.json"
CC_PROFILE_TXT = "CCProfile{}.txt"
ACCOUNT_TXT = "Account{}.txt"
UTF8 = "utf8"
SJIS = "sjis"
SATURDAY_INDEX = 5

CONFIG_TXT = "./config.txt"
CONFIG_DICT = {}
CONFIG_KEY_URL = "URL"


class ItemInfo:
    def __init__(self):
        self.index = None
        self.item_name = None
        self.keyword = None


class JsonMakerScreen(Screen):
    def __init__(self, **kwargs):
        super(JsonMakerScreen, self).__init__(**kwargs)
        self.item_list = []
        self.task_list_dict = {}
        self._file = Window.bind(on_dropfile=self._on_file_drop)
        self.excel_path = None
        self.dump_cnt = 0
        self.record_cnt = 0
        self.max_data_num_per_file = 0
        self.cc_profile_dict_list = []
        self.account_dict_list = []
        self.is_duplicate_profile = False
        self.did_finish_item_list_from_website = False
        self.did_finish_get_excel_path = False

    def _on_file_drop(self, window, file_path):
        self.disp_messg("{}を読み込みました".format(os.path.basename(file_path.decode("utf-8"))))
        self.excel_path = file_path.decode("utf-8")
        self.did_finish_get_excel_path = True
        if self.did_finish_item_list_from_website:
            self.ids[ID_DUMP_CHECKOUTPROFILES_BUTTON].disabled = False

    def disp_drag_and_drop_msg(self):
        self.disp_messg("Excelファイルをツール画面上にドラッグ&ドロップしてください")

    @staticmethod
    def get_latest_url():
        r = requests.get(CONFIG_DICT[CONFIG_KEY_URL], headers=HEADERS)
        soup = BeautifulSoup(r.text, "lxml")
        for a in soup.find_all("a"):
            if a.text == "Read More":
                return a.get("href")

    def get_site_info(self):
        self.disp_messg("サイトから最新情報を取得中...")
        self.ids[ID_GET_SITE_INFO_BUTTON].disabled = True
        Clock.schedule_once(self.update_item_view)

    def update_item_view(self, dt):
        try:
            url = self.get_latest_url()
            self.parse_site_info(url)
            scrollview = self.ids["container"]
            scrollview.clear_widgets()
            row_len = len(self.item_list)

            scrollview.height = row_len * 40
            for item_info in self.item_list:
                self.add_item_info_row(item_info, scrollview)

            self.disp_messg("サイトから最新情報を取得しました")
            self.did_finish_item_list_from_website = True

            if self.did_finish_get_excel_path:
                self.ids[ID_DUMP_CHECKOUTPROFILES_BUTTON].disabled = False

        except Exception as e:
            self.disp_messg_err("サイトから最新情報を取得するのに失敗しました。")
            log.exception("Unknown Exception : %s.", e)
        finally:
            self.ids[ID_GET_SITE_INFO_BUTTON].disabled = False

    def add_item_info_row(self, item_info, scrollview):
        box = BoxLayout()
        self.add_text_widget_on_grid(box, str(item_info.index), size_hint_x=0.15)
        self.add_text_widget_on_grid(box, item_info.item_name, size_hint_x=0.7)
        self.add_text_widget_on_grid(box, item_info.keyword, size_hint_x=0.7)
        scrollview.add_widget(box)

    @staticmethod
    def add_text_widget_on_grid(box, text, id=None, size_hint_x=1, disabled=True):
        text = TextInput(text=text, size_hint_x=size_hint_x, multiline=False, write_tab=False)
        text.id = id
        text.is_focusable = False
        # text.disabled = disabled
        text.disabled_foreground_color = (0, 0, 0, 1)
        text.background_disabled_normal = text.background_normal
        box.add_widget(text)

    def parse_site_info(self, url):
        self.item_list = []
        r = requests.get(url, headers=HEADERS)
        soup = BeautifulSoup(r.text, "lxml")

        index = 0
        for div in soup.find_all("div", class_="supreme product"):
            index += 1
            item_info = ItemInfo()
            item_info.index = index
            item_info.item_name = div.find("h3").text
            item_info.keyword = div.find("div", class_="supreme keywords").text
            self.item_list.append(item_info)

    def dump_json_files(self):
        try:
            self.dump_json_files_core()
        except Exception as e:
            self.disp_messg_err("ファイルの出力に失敗しました。")
            log.exception("ファイルの出力に失敗しました。%s", e)

    def dump_json_files_core(self):
        self.max_data_num_per_file = int(self.ids[ID_MAX_DATA_NUM_PER_FILE].text)
        self.dump_cnt = 0
        self.record_cnt = 0
        self.task_list_dict = {}
        self.cc_profile_dict_list = []
        workbook = open_workbook(self.excel_path)
        sheet = workbook.sheet_by_index(0)
        for i in range(1, sheet.nrows):
            row = sheet.row(i)
            log.info(row)

            if self.is_not_address_record(row):
                log.info("{}行目に必須項目未入力のセルがありました。この行の取り込みをスキップします".format(i + 1))
                continue

            if not self.mk_dict_list_and_dump(row, i + 1):
                return

        if 0 < len(self.cc_profile_dict_list):
            self.dump_2json_files()

        if self.dump_cnt > 1:
            self.disp_messg("{}行のレコードを{}分割してファイルに出力しました".format(self.record_cnt, self.dump_cnt))
        else:
            self.disp_messg("{}行のレコードをファイルに出力しました".format(self.record_cnt))

    def mk_dict_list_and_dump(self, row, row_num):
        item_no_list = self.split_list(row, INDEX_ITEM_NO)
        size_list = self.split_list(row, INDEX_ITEM_SIZE)

        if len(item_no_list) != len(size_list):
            self.disp_messg_err("{}行目のアイテム数とサイズの数が一致しません。出力を中断しました。".format(row_num))
            log.error("{}行目のアイテム数とサイズの数が一致しません。アイテム:{} サイズ:{}".format(
                row_num, item_no_list, size_list))
            return False

        self.is_duplicate_profile = False
        cc_profile_dict = self.mk_cc_profile_dict(row)

        for i in range(len(item_no_list)):

            if not self.is_duplicate_profile:
                self.cc_profile_dict_list.append(cc_profile_dict)
                self.is_duplicate_profile = True

            self.account_dict_list.append(self.mk_account_list(item_no_list[i], size_list[i], row))
            self.record_cnt += 1

            if self.max_data_num_per_file <= len(self.account_dict_list):
                self.dump_2json_files()

        return True

    def mk_account_list(self, item_no, size, row):
        account_dict = dict()
        account_dict[ACCOUNT_KEY_ADIDAS_EXPLOIT] = False
        account_dict[ACCOUNT_KEY_EMAIL] = row[INDEX_EMAIL].value
        account_dict[ACCOUNT_KEY_PASSWORD] = EMPTY
        account_dict[ACCOUNT_KEY_SIZE] = [size]

        if self.ids[ID_BUTTON_TYPE_PC].state == "down":
            account_dict[ACCOUNT_KEY_LINKS] = ["http://www.supremenewyork.com/shop/all"]
        else:
            account_dict[ACCOUNT_KEY_LINKS] = ["http://www.supremenewyork.com/mobile/"]

        account_dict[ACCOUNT_KEY_KEYWORDS] = [self.item_list[int(item_no) - 1].keyword]
        account_dict[ACCOUNT_KEY_NOTIFICATION_EMAIL] = EMPTY
        account_dict[ACCOUNT_KEY_NOTIFICATION_TEXT] = None

        if self.ids[ID_BUTTON_TYPE_PC].state == "down":
            account_dict[ACCOUNT_KEY_SITE_TYPE] = 5
        else:
            account_dict[ACCOUNT_KEY_SITE_TYPE] = 63

        account_dict[ACCOUNT_KEY_IS_GUEST] = True
        account_dict[ACCOUNT_KEY_DISABLED] = False

        checkout_info_dict = dict()
        checkout_info_dict[ACCOUNT_KEY_PAY_PAL_CHECKOUT] = False
        checkout_info_dict[ACCOUNT_KEY_CC_CHECKOUT] = True
        checkout_info_dict[ACCOUNT_KEY_FINALIZE_ORDER] = True
        checkout_info_dict[ACCOUNT_KEY_PAY_PAL_EMAIL_ADDRESS] = EMPTY
        checkout_info_dict[ACCOUNT_KEY_PAY_PAL_PASSWORD] = EMPTY
        checkout_info_dict[ACCOUNT_KEY_CC_PROFILE] = row[INDEX_PROFILE_NAME].value

        account_dict[ACCOUNT_KEY_CHECKOUT_INFO] = checkout_info_dict
        account_dict[ACCOUNT_KEY_ONLY_NEW_SUPREAME_PRODUCTS] = True
        account_dict[ACCOUNT_KEY_CHECKOUT_DELAY] = 0
        account_dict[ACCOUNT_KEY_QUANTITY] = 1

        return account_dict

    @staticmethod
    def mk_cc_profile_dict(row):
        cc_profile_dict = dict()
        cc_profile_dict[KEY_NAME] = row[INDEX_PROFILE_NAME].value
        for billing_or_shipping in BILLING_AND_SHIPPING:
            cc_profile_dict[billing_or_shipping + KEY_FIRST_NAME] = row[INDEX_FIRST_NAME].value
            cc_profile_dict[billing_or_shipping + KEY_LAST_NAME] = row[INDEX_LAST_NAME].value
            cc_profile_dict[billing_or_shipping + KEY_ADDRESS1] = row[INDEX_ADDRESS].value
            cc_profile_dict[billing_or_shipping + KEY_ADDRESS2] = EMPTY
            cc_profile_dict[billing_or_shipping + KEY_ZIP_CODE] = str(int(row[INDEX_ZIP_CODE].value))
            cc_profile_dict[billing_or_shipping + KEY_PHONE] = row[INDEX_PHONE].value
            cc_profile_dict[billing_or_shipping + KEY_CITY] = row[INDEX_CITY].value

            if billing_or_shipping == "Billing":
                cc_profile_dict[billing_or_shipping + KEY_STATE] = " " + row[INDEX_STATE].value
            else:
                cc_profile_dict[billing_or_shipping + KEY_STATE] = EMPTY

            cc_profile_dict[billing_or_shipping + KEY_PROVINCE] = EMPTY
            cc_profile_dict[billing_or_shipping + KEY_COUNTRY] = EMPTY
            cc_profile_dict[billing_or_shipping + KEY_EU_REGION] = EMPTY
            cc_profile_dict[billing_or_shipping + KEY_DSM_REGION] = EMPTY

        cc_profile_dict[KEY_CARD_TYPE] = row[INDEX_PAY_TYPE].value

        if row[INDEX_PAY_TYPE].value == "代金引換":
            cc_profile_dict[KEY_CARD_NUMBER] = EMPTY
            cc_profile_dict[KEY_CARD_EXPIRY_MONTH] = "1"
            cc_profile_dict[KEY_CARD_EXPIRY_YEAR] = "15"
            cc_profile_dict[KEY_CARD_CVV] = EMPTY
        else:
            cc_profile_dict[KEY_CARD_NUMBER] = row[INDEX_CARD_NUMBER].value
            cc_profile_dict[KEY_CARD_EXPIRY_MONTH] = str(int(row[INDEX_CARD_EXPIRY_MONTH].value))
            cc_profile_dict[KEY_CARD_EXPIRY_YEAR] = str(int(row[INDEX_CARD_EXPIRY_YEAR].value))
            cc_profile_dict[KEY_CARD_CVV] = str(int(row[INDEX_CARD_CVV].value))

        cc_profile_dict[KEY_MAX_CHECKOUTS] = 0
        return cc_profile_dict

    def dump_2json_files(self):
        self.dump_json_file(CC_PROFILE_TXT, self.cc_profile_dict_list)
        self.dump_json_file(ACCOUNT_TXT, self.account_dict_list)
        self.cc_profile_dict_list = []
        self.account_dict_list = []
        self.is_duplicate_profile = False
        self.dump_cnt += 1

    def dump_json_file(self, out_filename_fmt, out_list):
        if self.dump_cnt == 0:
            out_file_name = out_filename_fmt.format("")
        else:
            out_file_name = out_filename_fmt.format(self.dump_cnt)

        with open(out_file_name, "w", encoding=UTF8) as f:
            f.write(json.dumps(out_list, indent=2, ensure_ascii=False))

    def split_list(self, row, index):
        try:
            return row[index].value.split("&")
        except AttributeError:
            size = float(row[index].value)
            if size % 1.0 == 0.0:
                size = int(size)
            return [str(size)]

    @staticmethod
    def is_not_address_record(row):
        if row[INDEX_PROFILE_NAME].value == EMPTY \
                or row[INDEX_ITEM_NO].value == EMPTY or row[INDEX_ITEM_SIZE].value == EMPTY \
                or row[INDEX_FIRST_NAME].value == EMPTY or row[INDEX_LAST_NAME].value == EMPTY \
                or row[INDEX_ZIP_CODE].value == EMPTY or row[INDEX_STATE].value == EMPTY \
                or row[INDEX_CITY].value == EMPTY or row[INDEX_ADDRESS].value == EMPTY \
                or row[INDEX_PHONE].value == EMPTY or row[INDEX_PAY_TYPE].value == EMPTY \
                or row[INDEX_EMAIL].value == EMPTY:
            return True

        if row[INDEX_PAY_TYPE].value != "代金引換":
            if row[INDEX_CARD_NUMBER].value == EMPTY \
                    or row[INDEX_CARD_EXPIRY_MONTH].value == EMPTY \
                    or row[INDEX_CARD_EXPIRY_YEAR].value == EMPTY \
                    or row[INDEX_CARD_CVV].value == EMPTY:
                return True

        return False

    def press_btn_pc(self):
        self.ids[ID_BUTTON_TYPE_PC].state = "down"

    def press_btn_mobile(self):
        self.ids[ID_BUTTON_TYPE_MOBILE].state = "down"

    def disp_messg(self, msg):
        self.ids[ID_MESSAGE].text = msg
        self.ids[ID_MESSAGE].color = (0, 0, 0, 1)

    def disp_messg_err(self, msg):
        self.ids[ID_MESSAGE].text = "{}詳細はログファイルを確認してください。".format(msg)
        self.ids[ID_MESSAGE].color = (1, 0, 0, 1)


class JsonMakerApp(App):
    title = "大和型BNB Supreme補助ツール"

    def build(self):
        return JsonMakerScreen()


def match_license():
    return licensemanager.match_license()


def load_config():
    for line in open(CONFIG_TXT, "r", encoding=SJIS):
        items = line.replace("\n", "").split("=")

        if len(items) != 2:
            continue

        CONFIG_DICT[items[0]] = items[1]


if __name__ == '__main__':

    log = logging.getLogger('my-log')

    if not match_license():
        log.error("ライセンスエラー。プログラムを終了します。")
        sys.exit(-1)

    load_config()
    LabelBase.register(DEFAULT_FONT, "ipaexg.ttf")
    JsonMakerApp().run()
