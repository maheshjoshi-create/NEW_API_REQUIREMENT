# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 09:44:46 2021

@author: Swara_Kaivu
"""
import configparser
import regex as re
import pandas as pd
import numpy as np
from imap_tools import MailBox, AND, OR

def login_to_gmail():
    """This function is used to login to gmail.
    
    Args:
        None
        
    Returns:
        mailbox (obj): object of Gmail login
    
    """
    # Read the config.ini file
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
    from_list = ['submissions@forwardfinancing.com','mjackson@forwardfinancing.com']
    sender_list = [msg.from_ for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mails.fetch(AND(OR(from_=from_list),seen=False,subject='Application Incomplete for',all=True),mark_seen=False)]

    subjects_list = [' '.join(i.split()) for i in n_subjects]
    opportunity_name_list = [re.findall(r"(?<=Application Incomplete for )(.*)",i)[-1] for i in subjects_list]
    print("The length of the Forward Financing Add info email is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    ad_notes_list = []
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
        app_dec_notes = re.sub(r"(?<=<style>)[\S\s\n]*(?=</style>)",'',app_dec_notes,1)
        
        ad_notes_list.append(app_dec_notes)

    return opportunity_name_list,ad_notes_list, subjects_list, subjects_uid

def create_csv(o,adn,subjects,subjects_uid):
    """This function is used to create a csv file.
    
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
        "Status":'Additional Info Needed',
        "AdditionalInfoSTIPS__c":adn,
        "Approval Decline Notes":np.nan,
        "Subjects":subjects,
        "Subjects_uid":subjects_uid})

    df.to_csv("Forward_Financing_Add_Info.csv", index = False)


def main_forward_financing_add_info():
    """This function is used to extract the data from the email.
    
    Args:
        None
        
    Returns:
        None
    """
    inbox=login_to_gmail()
    o,adn,subjects_list, subjects_uid = mail_content(inbox)
    # mail_content(inbox)
    create_csv(o,adn,subjects_list, subjects_uid)

if __name__=="__main__":
    main_forward_financing_add_info()
