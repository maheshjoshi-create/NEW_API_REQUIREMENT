
import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND
import time

def login_to_gmail():
    # Read the config.ini file
    """ 
  Read the config.ini file using config = configparser.RawConfigParser()
   
  Login to Gmail using mailbox = mailbox.login
  """
    config=configparser.RawConfigParser()
    config.read('config.ini')

    # Login to Gmail
    mailbox = MailBox('imap.gmail.com')
    mailbox.login(config['GMAIL']['EMAIL'],config['GMAIL']['PASSWORD'], initial_folder='INBOX')

    return mailbox
def mail_content(mails):
    """This function is used to extract the data from the email.
    
    Args:
        mails: mailbox object
        
    Returns:
        opportunity_name_list (list): list of opportunity name,
        ad_notes_list (list): list of additional notes,
        subjects_list (list): list of subjects,
        subjects_uid (list): list of subjects uid
    """
    sender_list = [msg.from_ for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mails.fetch(AND(seen=False,subject='Application Ineligible for ',all=True),mark_seen=False)]

    subjects_list = [' '.join(i.split()) for i in n_subjects]
    subjects_list = [i.replace('\u200b','').replace('\x92','') for i in subjects_list]
    opportunity_name_list = [re.findall(r"(?<=Application Ineligible for ).*$",i)[-1] for i in subjects_list]
    ad_notes_list = []
    print("The Length of Forward_Financing_Decline Email is: ",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    for i in range(len(subjects_list)):
    # reciever mail extraction
        reciever_data = ""
        if type(reciever_list[i]) == tuple or type(reciever_list[i]) == list:
            for x in reciever_list[i]:
                reciever_data = reciever_data + x+", "
            reciever_data = reciever_data[:-2]
        else:
            reciever_data = reciever_list[i]
        # print(reciever_data)
        #content creations
        app_dec_notes = f"""<p>From: {sender_list[i]}</p>
        <p>Date: {date_list[i]}</p>
        <p>Subject: {subjects_list[i]}</p>
        <p>To: {reciever_data}</p>


        {mail_content[i]}
        """
        app_dec_notes = app_dec_notes.replace("p{margin-top:0px; margin-bottom:0px;}",'')
        ad_notes_list.append(app_dec_notes)

    return opportunity_name_list,ad_notes_list, subjects_list, subjects_uid

def create_csv(o,adn,subjects_list, subjects_uid):
    """This function is used to create csv file.
    
    Args:
        o (list): list of opportunity name,
        adn (list): list of additional notes,
        subjects (list): list of subjects,
        subjects_uid (list): list of subjects uid
        
    Returns:
        None
    """
    df = pd.DataFrame({"Opportunity":o,
        "Lender":'0014W00002CyCvaQAF',
        "Status":"Declined",
        "Approval Decline Notes":adn,
        "Subjects":subjects_list,
        "Subjects_uid":subjects_uid})
    df.to_csv("Forward_Financing_Decline.csv", index = False)


def main_forward_financing_decline():
    """This function is used to extract the data from the email.
    
    Args:
        None
        
    Returns:
        None
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    o,adn,subjects_list, subjects_uid = mail_content(inbox)
    create_csv(o,adn,subjects_list, subjects_uid)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Forward_Financing_Decline Time in minutes: ",execution_time//60)
    # for i in range(len(subjects_list)):
    #     # print(subjects_uid[i])
    #     pass
    # return subjects_list, subjects_uid
if __name__=="__main__":
    main_forward_financing_decline()