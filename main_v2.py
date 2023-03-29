from threading import Thread
import threading
from time import sleep
from ticketing_system_controls_v2 import ticketing_system_service_desk_browser
from exchange_controls_v2 import Ms_Exchange_Browser
from powershell_controls_v2 import AD_Group, AD_User
from request_viewer_controls_v2 import request_Browser
import win32com.client as win32
from os import getenv
import pythoncom
import logging
from main_secrets import *

logging.basicConfig(
    filename="logfile2.txt",
    level=logging.DEBUG,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%B %d %I:%M:%S%p",
)

logger = logging.getLogger()


def send_email(email_recipient, subject, body, cc_recipient=""):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = email_recipient
    mail.CC = cc_recipient
    mail.Subject = subject
    mail.Body = body
    # mail.Display(True)
    mail.Send()
    debug_body = body.replace("\n", "|")
    logger.debug(
        f"Email sent successfully with the following info. Email recipient={email_recipient}, subject={subject}, body={debug_body}, cc_recipient={cc_recipient}"
    )
    return True


current_ticket_queue = new_ticket_queue


def work_ticket(
    ticket_number,
    request_number,
    tickets_worked_counter,
    ticketing_system_browser,
    qr_browser,
    exchange_browser,
    thread_number,
):
    print(
        f"Thread {thread_number} | Starting work on ticketing_systemticket {ticket_number}, QR {request_number}"
    )
    ticket_information = ticketing_system_browser.parse_ticket(ticket_number)
    current_ticketing_system_ticket_summary = ticket_information[0]
    current_ticketing_system_ticket_description = ticket_information[1]
    logger.info(f"Thread {thread_number} | Opening request to review needed changes.")
    current_request = qr_browser.parse_request(request_number)
    current_distribution_list = AD_Group(current_request.dl_name)
    # Determine if shared mailbox
    if ("sendasaccess" in current_request.dl_name.lower()) or (
        "fullaccess" in current_request.dl_name.lower()
    ):
        ticket_action = "send email"
        prefilled_email_template = "shared_mailbox"
    # Determine if approval was not responded to
    elif (
        ("Owner not active" in current_ticketing_system_ticket_summary)
        or (
            "Inactive owner of Distribution List found."
            in current_ticketing_system_ticket_summary
        )
        or ("Owner not found" in current_ticketing_system_ticket_summary)
    ):
        logger.info(
            f"Thread {thread_number} | request had not validated approval, checking if requestor is an owner."
        )
        logger.info(
            f"Thread {thread_number} | Attempting to locate owners for {current_distribution_list.provided_name}"
        )
        print(f"Thread {thread_number} | Checking DL ownership.")
        current_distribution_list.query("Owners")
        if current_distribution_list.owners != None:
            if len(current_distribution_list.owners) == 0:
                current_distribution_list.owners = None
                current_distribution_list.query("Owners")
                if current_distribution_list.owners != None:
                    if len(current_distribution_list.owners) == 0:
                        pass
        ticket_action = None
        updated_owners = []
        if current_distribution_list.owners == None:
            logger.error(
                f"Thread {thread_number} | Unable to proceed, was unable to locate distribution list from provided name."
            )
            exit()
        for owner in current_distribution_list.owners:
            if (owner.employee_number).isnumeric() == False:
                logger.error(
                    f"Thread {thread_number} | Located incorrect owner, employeee number is showing: {owner.employee_number}"
                )
                exit()
            owner.query("mail")
            if owner.email != None:
                updated_owners.append(owner)
        current_distribution_list.owners = updated_owners
        for owner in current_distribution_list.owners:
            if owner.email == current_request.requestor_email:
                logger.info(f"Thread {thread_number} | Requestor is an owner of the DL")
                ticket_action = "process request"
                break
        if len(current_distribution_list.owners) == 0:
            logger.info(
                f"Thread {thread_number} | Detected no owners assigned to the DL"
            )
            ticket_action = "send email"
            prefilled_email_template = "no_owners"
        if ticket_action == None:
            ticket_action = "send email"
            prefilled_email_template = "no_approval_given"
            owners_email_list = []
            for owner in current_distribution_list.owners:
                owners_email_list.append(owner.email)
            logger.info(
                f"Thread {thread_number} | Current requestor {current_request.requestor_email} owners: {owners_email_list}"
            )
    else:
        ticket_action = "process request"
        email = current_distribution_list.query("email")
        if email == None:
            logger.error(
                f"Thread {thread_number} | Unable to locate email for {current_distribution_list.provided_name}."
            )
            exit()
    match ticket_action:
        case "pend ticket":
            owners_email_list = []
            for owner in current_distribution_list.owners:
                owners_email_list.append(owner.email)
            owner_emails = ", ".join(owners_email_list)
            update_notes = pend_ticket_template["update notes"] + "\n" + owner_emails
            logger.info(
                f"Thread {thread_number} | Opening ticketing_systemTicket to mark as pending"
            )
            ticketing_system_update_status = (
                ticketing_system_browser.update_ticket_status(
                    ticket_number,
                    new_ticket_status=pend_ticket_template["ticket status"],
                    new_group=pend_ticket_template["new group"],
                    new_asignee=pend_ticket_template["new asignee"],
                    update_notes=update_notes,
                )
            )
            while (attempts < 3) and (ticketing_system_update_status != True):
                attempts += 1
                ticketing_system_update_status = (
                    ticketing_system_browser.update_ticket_status(
                        ticket_number,
                        new_ticket_status=pend_ticket_template["ticket status"],
                        new_group=pend_ticket_template["new group"],
                        new_asignee=pend_ticket_template["new asignee"],
                        update_notes=update_notes,
                    )
                )
            if ticketing_system_update_status == False:
                logger.error(
                    f"Thread {thread_number} | Received error when attempting to pend ticket {ticket_number} in ticketing_system."
                )
                exit()
            else:
                logger.error(
                    f"Thread {thread_number} | Successfully pended ticket  {ticket_number}."
                )
                tickets_worked_counter.append(ticket_number)
        case "send email":
            print(f"Thread {thread_number} | Sending email")
            match prefilled_email_template:
                case "shared_mailbox":
                    email_status = send_email("OMITTED")
                    if email_status != True:
                        logger.error(
                            f"Thread {thread_number} | Error when attempting to send shared mailbox email for ticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Sent shared mailbox email for ticket {ticket_number}."
                        )
                    qr_status = qr_browser.abort_request(current_request.number)
                    attempts = 0
                    while (attempts < 3) and (qr_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to update request on attempt {attempts}. Re-attempting."
                        )
                        qr_status = qr_browser.abort_request(current_request.number)
                    if qr_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to abort request {current_request.number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully aborted request {current_request.number}."
                        )
                    ticketing_system_status = (
                        ticketing_system_browser.update_ticket_status(
                            ticket_number,
                            resolve_ticket_template_shared_mailbox["ticket status"],
                            resolve_ticket_template_shared_mailbox["new group"],
                            resolve_ticket_template_shared_mailbox["update notes"],
                            resolve_ticket_template_shared_mailbox["new asignee"],
                            resolve_ticket_template_shared_mailbox["resolution code"],
                        )
                    )
                    attempts = 0
                    while (attempts < 3) and (ticketing_system_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to  close ticketing_systemticket {ticket_number} on attempt {attempts}. Re-attempting."
                        )
                        ticketing_system_status = (
                            ticketing_system_browser.update_ticket_status(
                                ticket_number,
                                resolve_ticket_template_shared_mailbox["ticket status"],
                                resolve_ticket_template_shared_mailbox["new group"],
                                resolve_ticket_template_shared_mailbox["update notes"],
                                resolve_ticket_template_shared_mailbox["new asignee"],
                                resolve_ticket_template_shared_mailbox[
                                    "resolution code"
                                ],
                            )
                        )
                    if ticketing_system_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to close ticketing_systemticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully resolved ticket {ticket_number}."
                        )
                        tickets_worked_counter.append(ticket_number)
                case "no_approval_given":
                    email_status = send_email("OMITTED")
                    if email_status != True:
                        logger.error(
                            f"Thread {thread_number} | Error when attempting to send no approval email for ticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Sent no approval given email for {ticket_number}. Proceeding to terminate request."
                        )
                    qr_status = qr_browser.abort_request(current_request.number)
                    attempts = 0
                    while (attempts < 3) and (qr_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to update request on attempt {attempts}. Re-attempting."
                        )
                        qr_status = qr_browser.abort_request(current_request.number)
                    if qr_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to abort request {current_request.number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully aborted request {current_request.number}. Proceeding to close ticket."
                        )
                    ticketing_system_status = (
                        ticketing_system_browser.update_ticket_status(
                            ticket_number,
                            resolve_ticket_template_no_approval_given["ticket status"],
                            resolve_ticket_template_no_approval_given["new group"],
                            resolve_ticket_template_no_approval_given["update notes"],
                            resolve_ticket_template_no_approval_given["new asignee"],
                            resolve_ticket_template_no_approval_given[
                                "resolution code"
                            ],
                        )
                    )
                    while (attempts < 3) and (ticketing_system_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to  close ticketing_systemticket {ticket_number} on attempt {attempts}. Re-attempting."
                        )
                        ticketing_system_status = (
                            ticketing_system_browser.update_ticket_status(
                                ticket_number,
                                resolve_ticket_template_no_approval_given[
                                    "ticket status"
                                ],
                                resolve_ticket_template_no_approval_given["new group"],
                                resolve_ticket_template_no_approval_given[
                                    "update notes"
                                ],
                                resolve_ticket_template_no_approval_given[
                                    "new asignee"
                                ],
                                resolve_ticket_template_no_approval_given[
                                    "resolution code"
                                ],
                            )
                        )
                    if ticketing_system_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to close ticketing_systemticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully resolved ticket {ticket_number}."
                        )
                        tickets_worked_counter.append(ticket_number)
                case "no_owners":
                    email_status = send_email("OMITTED")
                    if email_status != True:
                        logger.error(
                            f"Thread {thread_number} | Error when attempting to send no owners email for ticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Sent no owners email for ticketing_systemticket {ticket_number}."
                        )
                    qr_status = qr_browser.abort_request(current_request.number)
                    attempts = 0
                    while (attempts < 3) and (qr_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to update request on attempt {attempts}. Re-attempting."
                        )
                        qr_status = qr_browser.abort_request(current_request.number)
                    if qr_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to abort request {current_request.number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully aborted request {current_request.number}."
                        )
                    ticketing_system_status = (
                        ticketing_system_browser.update_ticket_status(
                            ticket_number,
                            resolve_ticket_template_no_owners["ticket status"],
                            resolve_ticket_template_no_owners["new group"],
                            resolve_ticket_template_no_owners["update notes"],
                            resolve_ticket_template_no_owners["new asignee"],
                            resolve_ticket_template_no_owners["resolution code"],
                        )
                    )
                    while (attempts < 3) and (ticketing_system_status != True):
                        attempts += 1
                        logger.warning(
                            f"Thread {thread_number} | Recieved error when attempting to  close ticketing_systemticket {ticket_number} on attempt {attempts}. Re-attempting."
                        )
                        ticketing_system_status = (
                            ticketing_system_browser.update_ticket_status(
                                ticket_number,
                                resolve_ticket_template_no_owners["ticket status"],
                                resolve_ticket_template_no_owners["new group"],
                                resolve_ticket_template_no_owners["update notes"],
                                resolve_ticket_template_no_owners["new asignee"],
                                resolve_ticket_template_no_owners["resolution code"],
                            )
                        )
                    if ticketing_system_status != True:
                        logger.error(
                            f"Thread {thread_number} | Received error when attempting to close ticketing_systemticket {ticket_number}."
                        )
                        exit()
                    else:
                        logger.info(
                            f"Thread {thread_number} | Successfully resolved ticket {ticket_number}."
                        )
                        tickets_worked_counter.append(ticket_number)
        case "process request":
            updated_add_list = []
            updated_remove_list = []
            logger.info(f"Thread {thread_number} | Converting add list to emails.")
            for member in current_request.add_list:
                user = AD_User(member)
                user.query("mail")
                if user.email != None:
                    updated_add_list.append(user.email)
            current_request.add_list = updated_add_list
            if len(current_request.add_list) == 0:
                logger.info(f"Thread {thread_number} | No members to add.")
            logger.info(
                f"Thread {thread_number} | Converting remove list to display names."
            )
            for member in current_request.remove_list:
                user = AD_User(member)
                user.query("DisplayName")
                if user.display_name != None:
                    updated_remove_list.append(user.display_name)
            current_request.remove_list = updated_remove_list
            if len(current_request.remove_list) == 0:
                logger.info(f"Thread {thread_number} | No members to remove.")
            if len(current_request.add_list) > 0:
                logger.info(
                    f"Thread {thread_number} | Addtions to make: {current_request.add_list}"
                )
                print(f"Thread {thread_number} | Adding members to DL")
                add_status = exchange_browser.add_members(
                    current_distribution_list.email, current_request.add_list
                )
                attempts = 0
                while (attempts < 3) and (add_status != True):
                    attempts += 1
                    logger.warning(
                        f"Thread {thread_number} | Recieved error when attempting to add exchange members on attempt {attempts}. Re-attempting."
                    )
                    add_status = exchange_browser.add_members(
                        current_distribution_list.email, current_request.add_list
                    )
                if add_status != True:
                    if add_status == "DL permission error" or "DL config error":
                        update_notes = "Automation received error when attempting to modify members. Please process add/remove manually."
                        ticketing_system_status = (
                            ticketing_system_browser.update_ticket_status(
                                ticket_number,
                                new_ticket_status="Acknowledged",
                                new_group="MGTI GL ServiceDesk-Email",
                                new_asignee="zAutoAssignment, Analyst",
                                update_notes=update_notes,
                            )
                        )
                        if ticketing_system_status != True:
                            logger.error(
                                f"Thread {thread_number} | Received error when attempting to transfer ticketing_systemticket {ticket_number}."
                            )
                            exit()
                        else:
                            logger.info(
                                f"Thread {thread_number} | Successfully transferred ticket {ticket_number}."
                            )
                            tickets_worked_counter.append(ticket_number)
                            exit()
                    else:
                        logger.error(
                            f"Thread {thread_number} | Received error when adding members in Exchange to DL {current_distribution_list.email}."
                        )
                        exit()
                else:
                    logger.info(
                        f"Thread {thread_number} | Successfully added the following colleagues to {current_distribution_list.email}: {current_request.add_list}"
                    )
            else:
                logger.info(
                    f"Thread {thread_number} | No members to be added to DL, proceeding to removes."
                )
            if len(current_request.remove_list) > 0:
                logger.info(
                    f"Thread {thread_number} | Removals to make: {current_request.remove_list}"
                )
                print(f"Thread {thread_number} | Removing members from DL")
                remove_status = exchange_browser.remove_members(
                    current_distribution_list.email, current_request.remove_list
                )
                attempts = 0
                while (attempts < 3) and (remove_status != True):
                    attempts += 1
                    logger.warning(
                        f"Thread {thread_number} | Recieved error when attempting to remove exchange members on attempt {attempts}. Re-attempting."
                    )
                    remove_status = exchange_browser.remove_members(
                        current_distribution_list.email, current_request.remove_list
                    )
                if remove_status != True:
                    logger.error(
                        f"Thread {thread_number} | Received error when removing members in Exchangefrom DL {current_distribution_list.email}."
                    )
                    exit()
                else:
                    logger.info(
                        f"Thread {thread_number} | Successfully removed the following colleagues to {current_distribution_list.email}: {current_request.remove_list}"
                    )
            else:
                logger.info(
                    f"Thread {thread_number} | No members to be removed from DL, proceeding to abort request."
                )
            print(f"Thread {thread_number} | Aborting request")
            qr_status = qr_browser.abort_request(current_request.number)
            attempts = 0
            while (attempts < 3) and (qr_status != True):
                attempts += 1
                logger.warning(
                    f"Thread {thread_number} | Recieved error when attempting to update request on attempt {attempts}. Re-attempting."
                )
                qr_status = qr_browser.abort_request(current_request.number)
            if qr_status != True:
                logger.error(
                    f"Thread {thread_number} | Received error when attempting to abort request {current_request.number} after three attempts."
                )
                exit()
            else:
                logger.info(
                    f"Thread {thread_number} | Successfully aborted request {current_request.number}."
                )
            print(f"Thread {thread_number} | Resolving ticketing_systemTicket")
            ticketing_system_status = ticketing_system_browser.update_ticket_status(
                ticket_number,
                resolve_ticket_template_processed["ticket status"],
                resolve_ticket_template_processed["new group"],
                resolve_ticket_template_processed["update notes"],
                resolve_ticket_template_processed["new asignee"],
                resolve_ticket_template_processed["resolution code"],
            )
            attempts = 0
            while (attempts < 3) and (ticketing_system_status != True):
                attempts += 1
                logger.warning(
                    f"Thread {thread_number} | Recieved error when attempting to  close ticketing_systemticket {ticket_number} on attempt {attempts}. Re-attempting."
                )
                ticketing_system_status = ticketing_system_browser.update_ticket_status(
                    ticket_number,
                    resolve_ticket_template_processed["ticket status"],
                    resolve_ticket_template_processed["new group"],
                    resolve_ticket_template_processed["update notes"],
                    resolve_ticket_template_processed["new asignee"],
                    resolve_ticket_template_processed["resolution code"],
                )
            if ticketing_system_status != True:
                logger.error(
                    f"Thread {thread_number} | Received error when attempting to close ticketing_systemticket {ticket_number} after three attempts."
                )
                exit()
            else:
                logger.info(
                    f"Thread {thread_number} | Successfully resolved ticket {ticket_number}."
                )
                tickets_worked_counter.append(ticket_number)
    print(
        f"Thread {thread_number} | Finished working ticket {ticket_number}, request {request_number}"
    )
    logger.info(f"Finished working ticket {ticket_number}, request {request_number}")
    return


if __name__ == "__main__":
    chromedriver_path = "OMITTED"
    print("Starting intial browser loadup.")
    ticketing_system_browser1 = ticketing_system_service_desk_browser(chromedriver_path)
    qr_browser1 = request_Browser(chromedriver_path)
    exchange_browser1 = Ms_Exchange_Browser(
        chromedriver_path, getenv("EXCHANGE_USERNAME"), getenv("EXCHANGE_PASSWORD")
    )
    ticketing_system_browser2 = ticketing_system_service_desk_browser(chromedriver_path)
    qr_browser2 = request_Browser(chromedriver_path)
    exchange_browser2 = Ms_Exchange_Browser(
        chromedriver_path, getenv("EXCHANGE_USERNAME"), getenv("EXCHANGE_PASSWORD")
    )
    finished_status = False
    tickets_to_work = []
    tickets_being_worked = []
    tickets_finished = []
    thread_1_active = None
    thread_2_active = None
    logger.info("Browsers loaded, script is now starting.")
    print("Browsers loaded, script is now starting.")
    while finished_status == False:
        if tickets_to_work == False:
            logger.info(
                f"Finished working current queue, total of {tickets_finished} tickets processed."
            )
            finished_status = True
            break
        while len(tickets_to_work) < 2:
            if thread_1_active == None:
                tickets_to_work = ticketing_system_browser1.find_ticket_in_queue(
                    current_ticket_queue, tickets_to_work
                )
            else:
                tickets_to_work = ticketing_system_browser2.find_ticket_in_queue(
                    current_ticket_queue, tickets_to_work
                )
        for ticket in tickets_to_work:
            if ticket[0] not in tickets_being_worked:
                if thread_1_active == None:
                    tickets_being_worked.append(ticket[0])
                    my_thread = Thread(
                        target=work_ticket,
                        args=[
                            ticket[0],
                            ticket[1],
                            tickets_finished,
                            ticketing_system_browser1,
                            qr_browser1,
                            exchange_browser1,
                            "1",
                        ],
                    )
                    my_thread.start()
                    thread_1_active = ticket[0]
                    logger.info(f"Starting ticket {ticket[0]} on Thread 1")
                elif thread_2_active == None:
                    tickets_being_worked.append(ticket[0])
                    my_thread = Thread(
                        target=work_ticket,
                        args=[
                            ticket[0],
                            ticket[1],
                            tickets_finished,
                            ticketing_system_browser2,
                            qr_browser2,
                            exchange_browser2,
                            "2",
                        ],
                    )
                    my_thread.start()
                    thread_2_active = ticket[0]
                    logger.info(f"Starting ticket {ticket[0]} on Thread 2")
                else:
                    logger.critical("Unable to find empty thread.")
                    exit()
        while threading.active_count() > 2:
            sleep(1)
        new_tickets_to_work = []
        for ticket in tickets_to_work:
            if ticket[0] not in tickets_finished:
                new_tickets_to_work.append(ticket)
            else:
                if thread_1_active == ticket[0]:
                    thread_1_active = None
                elif thread_2_active == ticket[0]:
                    thread_2_active = None
                else:
                    logger.critical(
                        f"Unable to locate thread that was handling ticket. Current threads:{thread_1_active, thread_2_active}"
                    )
                    finished_status = True
                tickets_being_worked.remove(ticket[0])
        if len(new_tickets_to_work) == 2:
            logger.critical(
                "Thread closed without a ticket being processed. Terminating script"
            )
            finished_status = True
        tickets_to_work = new_tickets_to_work
        print(f"Tickets worked: {len(tickets_finished)}")
        logger.info(f"Tickets worked: {len(tickets_finished)}")
