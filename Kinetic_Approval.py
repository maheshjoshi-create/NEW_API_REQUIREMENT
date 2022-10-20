# -*- coding: utf-8 -*-
"""
Created on Wed Mar 16 15:19:47 2022

@author: user
"""

import configparser

import pandas as pd
import html2text
from imap_tools import MailBox, AND
import math
import time
import  numpy as np
import regex as re


def login_to_gmail():
    """This function is used to login to the gmail account and return the mailbox object.
    
    Args:
        None
        
    Returns:
        mailbox: MailBox object
    """
    
    config=configparser.RawConfigParser()
    config.read('config.ini')

    # Login to Gmail
    mailbox = MailBox('imap.gmail.com')
    mailbox.login(config['GMAIL']['EMAIL'],config['GMAIL']['PASSWORD'], initial_folder='INBOX')

    return mailbox
def mail_content(mailbox):
    """ This function extracts the data from the email and returns the data in the form of list.
    
    Args:
        mailbox: MailBox object
        
    Returns:
        opportunity_name_list (list): list of opportunity name,
        highest_funding_amount_list (list): list of highest funding amount,
        long_term (list): list of longest term,
        long_term_buyrate (list): list of longest term buyrate,
        ad_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        stips (list): list of funding stipulations
    """
    # list_ = ["stips@kdfunding.com"]
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(from_="stips@kdfunding.com",subject='Pre-Approval for '),seen = False, all=True),mark_seen=False)]
    subjects_list = [' '.join(i.split()) for i in n_subjects]
    
    high_funding_amt = [];long_term = [];ad_notes_list = [];stips = [];longterm_buy_rate=[]
    opportunity_name_list = [re.findall(r"(?<=Pre-Approval for )(.*)",i)[0] for i in subjects_list]
    print("_"*50)
    print("The length of Kinetic Approval is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    print("_"*50)
    for i in range(len(subjects_list)):
        # type(mail_content[i])
        df = pd.read_html(mail_content[i],match='Approved Amount')
        df=df[-1]
        # if "Term" not in df['Factor Rate'][0]:
        df.columns =[column.replace(" ", "_") for column in df.columns]
        amt = df['Approved_Amount'].iloc[0].replace('$','')
        vp = df['Points_Built_In'].iloc[0]/100
        
        amt = amt.replace(",","")
        high_funding_amt.append(amt)
        
        t = df['Term_Length'].iloc[0]
        long_term.append(t)
        
        buy = df['Factor_Rate'].iloc[0]
        br=buy-vp
        
        longterm_buy_rate.append(br)
        
        reciever_data = ""
        if type(reciever_list[i]) == tuple or type(reciever_list[i]) == list: 
            for x in reciever_list[i]: 
                reciever_data = reciever_data + x+", " 
            reciever_data = reciever_data[:-2]
        else: 
            reciever_data = reciever_list[i]


        app_dec_notes = f"""<p>From: {sender_list[i]}</p>
        <p>Date: {date_list[i]}</p>
        <p>Subject: {subjects_list[i]}</p>
        <p>To: {reciever_data}</p>


        {mail_content[i]}
        """
        app_dec_notes = re.sub(r"(?<=<head>)[\S\s\n]*(?=</head>)",'',app_dec_notes,1)
        ad_notes_list.append(app_dec_notes)
        
        doc = re.findall(r'(?<=<!--STIPS-->)(.*?)([\s\S]*)(?=<br><hr>)',mail_content[i])
        
        if doc ==[]:
            stips.append(np.nan)
        else:
            for a_tuple in doc:
                # print(i)
                stips.append(a_tuple[1])
        
    return opportunity_name_list,high_funding_amt,long_term,longterm_buy_rate,ad_notes_list,subjects_list, subjects_uid,stips

def create_csv(o,funding_amount,t1,longterm_buy_rate,adn,subject,subjects_uid,stips):
    """ This function creates a csv file and writes the data in the csv file.
    
    Args:
        o (list): list of opportunity name,
        funding_amount (list): list of highest funding amount,
        t1 (list): list of longest term,
        factor (list): list of longest term buyrate,
        adn (list): list of application decline notes,
        subject (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        stips (list): list of Stipulations
        
    Returns:
        None
    """
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002CxzH1QAJ',
        "status":"Approved",
        "highest funding amount":funding_amount,
        "longest term":t1,
        "longest term buyrate":longterm_buy_rate,
        "3 month buyrate":np.nan,
        "6 month buyrate":np.nan,
        "9 month buyrate":np.nan,
        "12 month buyrate":np.nan,
        "18 month buyrate":np.nan,
        "approval decline notes":adn,
        "Funding_Stipulations__c":stips,
        "subjects":subject,
        "subjects_uid":subjects_uid
        })
    df = df.dropna(subset=['approval decline notes'])
    df.to_csv("Kinetic_Approval.csv", index = False)
    
def main_kinetic_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        subject_list (list): list of subject of mail,
        subject_uid (list): list of uid of the mail
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    o,funding_amount,t1,longterm_buy_rate,adn,subjects_list,subjects_uid,stips= mail_content(inbox)
    create_csv(o,funding_amount,t1,longterm_buy_rate,adn,subjects_list, subjects_uid,stips)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Kinetic_Approval Time in minutes: ",execution_time//60)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_kinetic_approval()

