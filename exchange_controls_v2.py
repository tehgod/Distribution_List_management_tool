from lib2to3.pgen2 import driver
from time import sleep
from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    StaleElementReferenceException,
    TimeoutException,
    NoSuchElementException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
import logging
from selenium.webdriver.remote.remote_connection import LOGGER as seleniumLogger
from urllib3.connectionpool import log as urllibLogger
from os import getenv


urllibLogger.setLevel(logging.WARNING)
seleniumLogger.setLevel(logging.WARNING)

logger = logging.getLogger(__name__)

load_dotenv()


class Ms_Exchange_Browser:
    def __init__(self, driver_path, exchange_username, exchange_password):
        self.exchange_username = exchange_username
        self.exchange_password = exchange_password
        ca_ticket_browser_options = Options()
        ca_ticket_browser_options.add_argument("--window-size=1920,1080")
        ca_ticket_browser_options.add_argument("--log-level=3")
        ca_ticket_browser_options.add_argument("--headless")
        ca_ticket_browser_options.add_argument("--disable-logging")
        self.driver = webdriver.Chrome(
            executable_path=driver_path, options=ca_ticket_browser_options
        )
        logger.debug("Exchange browser loaded.")

    def add_members(self, distribution_list, list_of_members_emails):
        wait = WebDriverWait(self.driver, 60)
        if self.login() == False:
            logger.warning("Failed to login when adding members.")
            return False
        wait.until(EC.element_to_be_clickable((By.ID, "Menu_DistributionGroups")))
        self.driver.find_element(By.ID, "Menu_DistributionGroups").click()
        sleep(3)
        self.driver.switch_to.frame(0)
        sleep(3)
        wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "//*[@id='ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar']/div[13]/a",
                )
            )
        )
        self.driver.find_element(
            By.XPATH,
            "//*[@id='ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar']/div[13]/a",
        ).click()
        wait.until(
            EC.visibility_of_element_located(
                (
                    By.ID,
                    "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_SearchBox",
                )
            )
        )
        self.driver.find_element(
            By.ID,
            "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_SearchBox",
        ).send_keys(distribution_list)
        self.driver.find_element(
            By.ID,
            "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_SearchBox_SearchButton",
        ).click()
        sleep(2)
        try:
            self.driver.find_element(
                By.ID,
                "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_contentTableParent",
            ).click()
            self.driver.find_element(
                By.XPATH,
                '//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar"]/div[3]/a',
            ).click()
        except:
            try:
                self.driver.find_element(
                    By.ID,
                    "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_contentTableParent",
                ).click()
                self.driver.find_element(
                    By.XPATH,
                    '//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar"]/div[5]/a',
                ).click()
            except:
                logger.warning("Unable to locate DL in MS Exchange")
                return False
        window_found = False
        attempts = 0
        wait = WebDriverWait(self.driver, 30)
        while attempts < 3:
            try:
                attempts += 1
                wait.until(EC.number_of_windows_to_be(2))
                window_found = True
            except:
                logger.debug("Failed to load second window, retrying.")
        if window_found == False:
            logger.warning(
                "Unable to switch to DL properties window, terminating script."
            )
            return False
        wait = WebDriverWait(self.driver, 120)
        sleep(2)
        second_window = self.driver.window_handles[1]
        main_window = self.driver.window_handles[0]
        self.driver.switch_to.window(second_window)
        wait.until(
            EC.presence_of_element_located(
                (By.ID, "ResultPanePlaceHolder_caption_textContainer")
            )
        )
        wait.until(EC.element_to_be_clickable, ((By.ID, "bookmarklink_2")))
        if len(self.driver.find_elements(By.ID, "dlgModalError_tdDlgBdy")) > 0:
            logger.warning(
                f"Distribution List {distribution_list} has pending errors that need corrected first, routing back to service desk for manual fixing."
            )
            return "DL config error"
        self.driver.find_element(By.ID, "bookmarklink_2").click()
        sleep(2)
        self.driver.find_element(
            By.XPATH,
            '//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_ToolBar"]/div[1]/a',
        ).click()
        wait.until(EC.number_of_windows_to_be(3))
        third_window = self.driver.window_handles[2]
        self.driver.switch_to.window(third_window)
        wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="ResultPanePlaceHolder_pickerContent_pickerListView_ToolBar"]/div[1]/a',
                )
            )
        )
        sleep(2)
        self.driver.find_element(
            By.XPATH,
            '//*[@id="ResultPanePlaceHolder_pickerContent_pickerListView_ToolBar"]/div[1]/a',
        ).click()
        for email_address in list_of_members_emails:
            try:
                self.driver.find_element(
                    By.NAME,
                    "ctl00$ResultPanePlaceHolder$pickerContent$pickerListView$SearchBox",
                ).clear()
                self.driver.find_element(
                    By.NAME,
                    "ctl00$ResultPanePlaceHolder$pickerContent$pickerListView$SearchBox",
                ).send_keys(email_address)
                self.driver.find_element(
                    By.NAME,
                    "ctl00$ResultPanePlaceHolder$pickerContent$pickerListView$SearchBox",
                ).send_keys(Keys.ENTER)
                sleep(1)
                self.driver.find_element(
                    By.ID,
                    "ResultPanePlaceHolder_pickerContent_pickerListView_contentTableParent",
                ).click()
                self.driver.find_element(
                    By.ID, "ResultPanePlaceHolder_pickerContent_btnAddItem"
                ).click()
                self.driver.find_element(
                    By.NAME,
                    "ctl00$ResultPanePlaceHolder$pickerContent$pickerListView$SearchBox",
                ).clear()
            except:
                logger.warning(
                    f"Error when attempting to add email {email_address}, proceeding to next email."
                )
        sleep(3)
        self.driver.find_element(
            By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCommit"
        ).click()
        sleep(3)
        self.driver.switch_to.window(second_window)
        self.driver.find_element(
            By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCommit"
        ).click()
        sleep(2)
        try:
            error_body = ""
            error_body = self.driver.find_elements(By.ID, "dlgModalError_tdDlgBdy")
        except:
            pass
        if len(error_body) > 0:
            logger.debug(
                "Distribution List has permission errors, routing back to service desk for manual fixing."
            )
            self.driver.find_element(By.ID, "dlgModalError_OK").click()
            self.driver.find_element(
                By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCancel"
            ).click()
            return "DL permission error"
        self.driver.switch_to.window(main_window)
        logger.debug("Add_members function finished successfully.")
        return True

    def remove_members(self, distribution_list, list_of_members_names):
        wait = WebDriverWait(self.driver, 60)
        if self.login() == False:
            logger.warning("Failed to login when removing members.")
            return False
        wait.until(EC.element_to_be_clickable((By.ID, "Menu_DistributionGroups")))
        self.driver.find_element(By.ID, "Menu_DistributionGroups").click()
        sleep(3)
        self.driver.switch_to.frame(0)
        self.driver.find_element(
            By.XPATH,
            "//*[@id='ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar']/div[13]/a",
        ).click()
        self.driver.find_element(
            By.ID,
            "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_SearchBox",
        ).send_keys(distribution_list)
        self.driver.find_element(
            By.ID,
            "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_SearchBox_SearchButton",
        ).click()
        sleep(2)
        try:
            self.driver.find_element(
                By.ID,
                "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_contentTableParent",
            ).click()
            self.driver.find_element(
                By.XPATH,
                '//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar"]/div[3]/a',
            ).click()
        except:
            try:
                self.driver.find_element(
                    By.ID,
                    "ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_contentTableParent",
                ).click()
                self.driver.find_element(
                    By.XPATH,
                    '//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar"]/div[5]/a',
                ).click()
            except:
                logger.warning(
                    f"Unable to locate DL {distribution_list} in MS Exchange"
                )
                return False
        window_found = False
        attempts = 0
        wait = WebDriverWait(self.driver, 30)
        while attempts < 3:
            try:
                attempts += 1
                wait.until(EC.number_of_windows_to_be(2))
                window_found = True
            except:
                logger.debug("Failed to load second window, retrying.")
        if window_found == False:
            logger.warning(
                "Unable to switch to DL properties window, terminating script."
            )
            return False
        wait.until(EC.number_of_windows_to_be(2))
        wait = WebDriverWait(self.driver, 60)
        sleep(2)
        second_window = self.driver.window_handles[1]
        main_window = self.driver.window_handles[0]
        self.driver.switch_to.window(second_window)
        wait.until(
            EC.presence_of_element_located(
                (By.ID, "ResultPanePlaceHolder_caption_textContainer")
            )
        )
        self.driver.find_element(By.ID, "bookmarklink_2").click()
        sleep(2)
        for name in list_of_members_names:
            member_list_raw = self.driver.find_element(
                By.ID,
                "ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_contentTable",
            ).get_attribute("innerHTML")
            members_list = []
            current_position = 0
            list_done = False
            while list_done == False:
                current_name_start = (
                    member_list_raw.find("title=", current_position) + 7
                )
                current_name_end = member_list_raw.find('"', current_name_start)
                if current_name_start == 6:
                    list_done = True
                else:
                    members_list.append(
                        member_list_raw[current_name_start:current_name_end]
                    )
                current_position = current_name_end
            sleep(1)
            try:
                name_location = members_list.index(name) + 1
                self.driver.find_element(
                    By.XPATH,
                    f'//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_contentTable"]/tbody/tr[{name_location}]/td',
                ).click()
                sleep(1)
                self.driver.find_element(
                    By.XPATH,
                    '//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_ToolBar"]/div[3]/a',
                ).click()
                sleep(1)
                logger.debug(
                    f"Successfully removed {name} from DL, proceeding to next name"
                )
            except:
                try:
                    for current_name in members_list:
                        if current_name.startswith(name):
                            try:
                                name_location = members_list.index(current_name) + 1
                                self.driver.find_element(
                                    By.XPATH,
                                    f'//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_contentTable"]/tbody/tr[{name_location}]/td',
                                ).click()
                                sleep(1)
                                self.driver.find_element(
                                    By.XPATH,
                                    '//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_ToolBar"]/div[3]/a',
                                ).click()
                                sleep(1)
                                logger.debug(
                                    f"Successfully removed {name} from DL, proceeding to next name"
                                )
                            except:
                                logger.debug(
                                    f"{name} not located in current member list, proceeding to next name"
                                )
                except:
                    logger.debug(
                        f"{name} not located in current member list, proceeding to next name"
                    )
                    sleep(1)
        self.driver.find_element(
            By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCommit"
        ).click()
        sleep(2)
        try:
            error_body = ""
            error_body = self.driver.find_elements(By.ID, "dlgModalError_tdDlgBdy")
        except:
            pass
        if len(error_body) > 0:
            logger.debug(
                "Distribution List has permission errors, routing back to service desk for manual fixing."
            )
            self.driver.find_element(By.ID, "dlgModalError_OK").click()
            self.driver.find_element(
                By.ID, "ResultPanePlaceHolder_ButtonsPanel_btnCancel"
            ).click()
            return "DL permission error"
        self.driver.switch_to.window(main_window)
        logger.debug("remove_members function finished successfully.")
        return True

    def login(self):
        self.driver.get("OMITTED")
        try:
            self.driver.find_element(By.ID, "Menu_DistributionGroups")
            logger.debug("Exchange detected to already be logged in at menu screen.")
            return True
        except:
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "username"))
                )
                self.driver.find_element(By.ID, "username").send_keys(
                    self.exchange_username
                )
                self.driver.find_element(By.ID, "password").send_keys(
                    self.exchange_password
                )
                self.driver.find_element(
                    By.XPATH, "//*[@id='lgnDiv']/div[9]/div"
                ).click()
                sleep(5)
                try:
                    loading_error = self.driver.find_element(By.ID, "signInErrorDiv")
                    print(
                        f"""Recieved error "{loading_error.text}" when logging into Exchange."""
                    )
                    logger.warning(
                        f"""Recieved error "{loading_error.text}" when logging into Exchange."""
                    )
                    return False
                except NoSuchElementException:  # expected, means login successful
                    pass
                try:
                    self.driver.find_element(By.ID, "Menu_DistributionGroups")
                    logger.debug(
                        "Successfully signed into Exchange and detected menu screen."
                    )
                    return True
                except:
                    logger.error(
                        "Login process failed for Exchange for unknown reason."
                    )
                    return False
            except:
                logger.warning("Neither login screen or main menu detected.")
                return False


if __name__ == "__main__":
    browser = Ms_Exchange_Browser(
        "chromedriver.exe", getenv("EXCHANGE_USERNAME"), getenv("EXCHANGE_PASSWORD")
    )
