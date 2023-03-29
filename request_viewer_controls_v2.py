from os import abort
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
from time import sleep
import logging
from selenium.webdriver.remote.remote_connection import LOGGER as seleniumLogger
from urllib3.connectionpool import log as urllibLogger

urllibLogger.setLevel(logging.WARNING)
seleniumLogger.setLevel(logging.WARNING)

logger = logging.getLogger(__name__)


class request:
    def __init__(
        self,
        qr_number,
        requestor_name,
        requestor_empid,
        requestor_email,
        dl_name,
        dl_changes,
        qr_status,
        ticketing_system_ticket_number,
    ):
        self.number = qr_number
        self.requestor_name = requestor_name
        self.requestor_empid = requestor_empid
        self.requestor_email = requestor_email
        self.dl_name = dl_name
        self.dl_changes = dl_changes
        self.status = qr_status
        self.remove_list = []
        self.add_list = []
        self.ticketing_system_ticket_number = ticketing_system_ticket_number
        for change in self.dl_changes:
            if change.find("@") == -1:
                employee_id = change[change.find(" ") + 1 :]
            else:
                if change.startswith("Add"):
                    employee_id = change[change.find(" ", 4) + 1 :]
                elif change.startswith("Remove"):
                    employee_id = change[change.find(" ", 7) + 1 :]
            if change.startswith("Add"):
                if employee_id not in self.add_list:
                    self.add_list.append(employee_id)
            elif change.startswith("Remove"):
                if employee_id not in self.remove_list:
                    self.remove_list.append(employee_id)
        logger.debug(f"Created instance request class.")


class request_Browser:
    def __init__(self, driver_path):
        request_browser_options = Options()
        request_browser_options.add_argument("--window-size=1920,1080")
        request_browser_options.add_argument("--log-level=3")
        request_browser_options.add_argument("--headless")
        request_browser_options.add_argument("--disable-logging")
        self.driver = webdriver.Chrome(
            executable_path=driver_path, options=request_browser_options
        )
        logger.debug("QR browser loaded.")

    def parse_request(self, full_qr_number):
        wait = WebDriverWait(self.driver, 60)
        attempts = 0
        request_loaded = False
        while attempts < 3:
            try:
                attempts += 1
                if full_qr_number.startswith("1"):
                    self.driver.get(f"OMITTED{full_qr_number[:7]}OMITTED")
                else:
                    self.driver.get(f"OMITTED")
                wait.until(
                    EC.visibility_of_element_located(
                        (By.ID, f"heading{full_qr_number}")
                    )
                )
                attempts = 3
                request_loaded = True
                logger.debug(
                    f"Successfully loaded request page for request {full_qr_number}."
                )
            except:
                logger.debug(
                    f"Failed to load request page, trying again. Attempt {attempts}"
                )
        if request_loaded != True:
            logger.warning(
                "failed to load request after three attempts, terminating script."
            )
            return False
        WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.ID, f"collapse{full_qr_number}"))
        )
        qr_status = self.driver.find_element(By.ID, f"collapse{full_qr_number}").text
        WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[2]/div[2]',
                )
            )
        )
        requestor_information = self.driver.find_element(
            By.XPATH,
            f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[2]/div[2]',
        ).text
        requestor_information = requestor_information.split("\n")
        requestor_empid = requestor_information[0][12:]
        requestor_name = (
            f"{requestor_information[3][10:]}, {requestor_information[2][11:]}"
        )
        requestor_email = (requestor_information[6][14:]).lower()
        dl_name = self.driver.find_element(
            By.XPATH,
            f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[5]/div[2]/table/tbody/tr/td[2]',
        ).text
        if "\\" in dl_name:
            end_pos = dl_name.find("\\")
            dl_name = dl_name[:end_pos]
        try:
            dl_changes = self.driver.find_element(
                By.XPATH,
                f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[11]/div[2]',
            ).text
            dl_changes = dl_changes.split("\n")
            dl_changes.remove(dl_changes[0])
        except NoSuchElementException:
            dl_change_action = self.driver.find_element(
                By.XPATH,
                f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[9]/div[2]',
            ).text
            dl_change_target = self.driver.find_element(
                By.XPATH,
                f'//*[@id="tbl_data_{full_qr_number}"]/div/div[1]/div/div/div/div[2]/div/div[10]/div[2]/table/tbody/tr[1]/td[2]',
            ).text
            dl_changes = [f"{dl_change_action} {dl_change_target}"]
        qr_history = (
            self.driver.find_element(By.CSS_SELECTOR, ".list-unstyled.timeline")
        ).find_elements(By.CSS_SELECTOR, ".activity")
        qr_history_last_update = qr_history[len(qr_history) - 1].text
        if "ticketing_system Error  ticket #" in qr_history_last_update:
            ticket_number = qr_history_last_update[
                qr_history_last_update.find("#")
                + 1 : qr_history_last_update.find("#")
                + 9
            ]
        else:
            ticket_number = None
        logger.debug(f"Finished running parse request on QR {full_qr_number}")
        return request(
            full_qr_number,
            requestor_name,
            requestor_empid,
            requestor_email,
            dl_name,
            dl_changes,
            qr_status,
            ticket_number,
        )

    def abort_request(self, full_qr_number):
        wait = WebDriverWait(self.driver, 120)
        if full_qr_number.startswith("1"):
            self.driver.get(f"OMITTED")
        else:
            self.driver.get(f"OMITTED")
        wait.until(
            EC.visibility_of_element_located((By.ID, f"heading{full_qr_number}"))
        )
        sleep(3)
        current_status = self.driver.find_element(
            By.XPATH, f'//*[@id="tbl_data_{full_qr_number}"]/div/div[2]/dl/dd[1]'
        ).text
        if current_status.endswith("(Aborted)"):
            logger.debug(f"request {full_qr_number} detected to already be aborted.")
            return True
        try:
            abort_button = self.driver.find_element(
                By.XPATH, f'//*[@id="tbl_data_{full_qr_number}"]/div/div[2]/dl/dd[3]/a'
            )
        except NoSuchElementException:
            qr_inner_status = self.driver.find_element(
                By.XPATH, '//*[@id="subContainer"]/div[3]/div/div[2]/span'
            )
            if qr_inner_status.text == "Completed":
                logger.debug(
                    f"request {full_qr_number} detected to already be aborted."
                )
                return True
            else:
                logger.debug(
                    f"request {full_qr_number} missing abort button, not showing as completed. Instead showing as {qr_inner_status}"
                )
                return False
        abort_button.click()
        abort_status = False
        attempts = 0
        while (abort_status == False) or (attempts < 20):
            attempts += 1
            current_status = self.driver.find_element(
                By.XPATH, f'//*[@id="tbl_data_{full_qr_number}"]/div/div[2]/dl/dd[1]'
            ).text
            if current_status.endswith("(Aborted)"):
                abort_status = True
            else:
                logger.debug(
                    f"Status on request not updated to aborted on attempt {attempts}"
                )
                sleep(1)
        if abort_status != True:
            logger.warning(
                f"Failed to locate aborted status on qr {full_qr_number} after 20 attempts. Resolving out."
            )
            return False
        logger.debug(f"Successfully aborted request {full_qr_number}")
        return True


if __name__ == "__main__":
    pass
