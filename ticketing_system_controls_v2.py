from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep
import logging
from selenium.webdriver.remote.remote_connection import LOGGER as seleniumLogger
from urllib3.connectionpool import log as urllibLogger

urllibLogger.setLevel(logging.WARNING)
seleniumLogger.setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


class ticketing_system_service_desk_browser:
    def __init__(self, driver_path):
        ticketing_system_ticket_browser_options = Options()
        ticketing_system_ticket_browser_options.add_argument("--window-size=1920,1080")
        ticketing_system_ticket_browser_options.add_argument("--log-level=3")
        ticketing_system_ticket_browser_options.add_argument("--disable-logging")
        ticketing_system_ticket_browser_options.add_argument("--headless")
        self.driver = webdriver.Chrome(
            executable_path=driver_path, options=ticketing_system_ticket_browser_options
        )
        logger.debug("ticketing_system browser loaded.")

    def find_ticket_in_queue(self, queue_url, ticket_exclusion_list=[]):
        wait = WebDriverWait(self.driver, 30)
        self.driver.get(queue_url)
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
            wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "role_main"))
            )
        except TimeoutException:
            logger.warning("Failed to load webpage.")
            return False
        try:
            ticket_number = self.driver.find_element(By.ID, "1").text[:8]
            quickrequest_number = self.driver.find_element(By.ID, "sub_1_summary").text[
                29:44
            ]
            if quickrequest_number.startswith("1"):
                quickrequest_number = self.driver.find_element(
                    By.ID, "sub_1_summary"
                ).text[29:45]
            located_ticket = [ticket_number, quickrequest_number]
            if located_ticket in ticket_exclusion_list:
                ticket_number = self.driver.find_element(By.ID, "2").text[:8]
                quickrequest_number = self.driver.find_element(
                    By.ID, "sub_2_summary"
                ).text[29:44]
                if quickrequest_number.startswith("1"):
                    quickrequest_number = self.driver.find_element(
                        By.ID, "sub_2_summary"
                    ).text[29:45]
                located_ticket = [ticket_number, quickrequest_number]
            ticket_exclusion_list.append(located_ticket)
            logger.debug("find_ticket_in_queue finished successfully.")
            return ticket_exclusion_list
        except NoSuchElementException:
            logger.warning("Error locating next ticket")
            return False

    def update_ticket_status(
        self,
        ticket_number,
        new_ticket_status,
        new_group,
        update_notes,
        new_asignee=None,
        resolution_code=None,
    ):
        wait = WebDriverWait(self.driver, 60)
        self.driver.get(f"OMITTED{ticket_number}")
        self.driver.switch_to.default_content()
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "menubar")))
        except TimeoutException:
            logger.warning(
                f"Failed to load ticketing system page for ticket {ticket_number}."
            )
            return False
        self.driver.find_element_by_id("menu_2").click()  # Click the activities button
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame("product")
        self.driver.switch_to.frame("role_main")
        if (self.driver.find_element(By.ID, "df_0_2").text) == new_ticket_status:
            logger.debug("Ticket detected to be in correct state already.")
            return True
        self.driver.find_element_by_id("amActivities_1").click()
        logger.debug("Selected update status button")
        wait.until(EC.number_of_windows_to_be(2))
        sleep(3)
        second_window = self.driver.window_handles[1]
        main_window = self.driver.window_handles[0]
        sleep(1)
        self.driver.switch_to.window(second_window)
        logger.debug("Located second window appeared.")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "cai_main")))
        sleep(3)
        wait.until(
            EC.presence_of_element_located((By.ID, "df_1_1"))
        )  # status update window loaded
        Select(self.driver.find_element(By.ID, "df_1_1")).select_by_visible_text(
            new_ticket_status
        )  # set status
        if resolution_code != None:
            self.driver.find_element(By.ID, "df_3_1").send_keys(resolution_code)
        current_group_field = self.driver.find_element(By.ID, "df_5_0")
        current_group_field.clear()
        current_assignee_field = self.driver.find_element(By.ID, "df_5_1")
        current_assignee_field.clear()
        current_group_field.send_keys(new_group)  # set resolver group
        if new_asignee != None:
            current_assignee_field.send_keys(new_asignee)  # set asignee
        self.driver.find_element_by_id("df_7_0").send_keys(
            update_notes
        )  # set resolution notes
        sleep(3)
        self.driver.find_element_by_id("imgBtn0").click()
        try:
            wait.until(EC.number_of_windows_to_be(1))
        except TimeoutException:
            logger.warning("Unable to detect if update status window closed out.")
        sleep(3)
        self.driver.switch_to.window(main_window)
        try:
            self.driver.switch_to.default_content()
        except TimeoutException:
            return
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "role_main")))
        resolved_status = False
        attempts = 0
        while (attempts < 5) and (resolved_status != True):
            msg_bar = self.driver.find_element(By.ID, "alertmsgText").text
            if "Save Successful" in msg_bar:
                logger.info("Save successful located on main ticket window")
                resolved_status = True
            else:
                sleep(5)
                logger.warning(
                    f"Was unable to detect save status on ticket during attempt {attempts}."
                )
                attempts += 1
        if resolved_status == False:
            logger.warning(f"Failed to detect second window after fiive attempts.")
            return False
        else:
            logger.debug("Update_ticket_status function returning true.")
            return True

    def transfer_ticket(
        self, ticket_number, new_group, transfer_notes, new_asignee=None
    ):
        wait = WebDriverWait(self.driver, 60)
        self.driver.get(f"OMITTED{ticket_number}")
        self.driver.switch_to.default_content()
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "menubar")))
        except TimeoutException:
            logger.warning(f"Unable to load ticket {ticket_number}.")
            return False
        self.driver.find_element_by_id("menu_2").click()  # Click the activities button
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame("product")
        self.driver.switch_to.frame("role_main")
        if (self.driver.find_element(By.ID, "df_0_2").text) == "Resolved":
            logger.debug("Ticket is detected to be already resolved.")
            return True
        self.driver.find_element_by_id(
            "amActivities_7"
        ).click()  # click transfer button
        wait.until(EC.number_of_windows_to_be(2))
        sleep(2)
        second_window = self.driver.window_handles[1]
        main_window = self.driver.window_handles[0]
        self.driver.switch_to.window(second_window)
        logger.debug("Located second window appeared.")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "cai_main")))
        wait.until(
            EC.presence_of_element_located((By.ID, "df_4_1"))
        )  # transfer window loaded
        current_assignee_field = self.driver.find_element(By.ID, "df_4_1")
        current_assignee_field.clear()
        new_group_field = self.driver.find_element(By.ID, "df_4_0")
        new_group_field.clear()
        new_group_field.send_keys(new_group)
        if new_asignee != None:
            current_assignee_field.send_keys(new_asignee)
        self.driver.find_element(By.ID, "df_6_0").send_keys(transfer_notes)
        sleep(1)
        self.driver.find_element(By.ID, "imgBtn0").click()
        wait.until(EC.number_of_windows_to_be(1))
        self.driver.switch_to.window(main_window)
        self.driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "role_main")))
        resolved_status = False
        attempts = 0
        while (attempts < 5) and (resolved_status != True):
            msg_bar = self.driver.find_element(By.ID, "alertmsgText").text
            if "Save Successful" in msg_bar:
                logger.info("Save successful located on main ticket window")
                resolved_status = True
            else:
                sleep(5)
                logger.warning(
                    f"Was unable to detect save status on ticket during attempt {attempts}."
                )
                attempts += 1
        if resolved_status == False:
            logger.warning(f"Failed to detect second window after fiive attempts.")
            return False
        else:
            return True

    def parse_ticket(self, ticket_number):
        wait = WebDriverWait(self.driver, 60)
        self.driver.get(f"OMITTED{ticket_number}")
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "product")))
            wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "role_main"))
            )
        except TimeoutException:
            logger.warning(f"Failed to load webpage for ticket number {ticket_number}")
            return False
        ticket_summary = self.driver.find_element(By.ID, "df_11_0").text
        ticket_description = self.driver.find_element(By.ID, "df_12_0").text
        logger.debug(f"Successfully parsed ticket {ticket_number}")
        return [ticket_summary, ticket_description]


if __name__ == "__main__":
    my_browser = ticketing_system_service_desk_browser("C:/webdriver/chromedriver.exe")
    ticket_queue = "OMITTED"
    my_queue = my_browser.find_ticket_in_queue(ticket_queue, my_queue)
    print(my_queue)
    my_browser.find_ticket_in_queue(ticket_queue, my_queue)
    print(my_queue)
