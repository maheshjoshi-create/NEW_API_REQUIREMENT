# -*- coding: utf-8 -*-
"""
Created on Tue Jul 26 10:49:35 2022

@author: user
"""


import time
import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND
import time

import numpy as np

def login_to_gmail():
    """This function is used to login to the gmail account and return the mailbox object.
    
    Args:
        None
        
    Returns:
        mailbox: MailBox object
    """
    # Read the config.ini file
    config=configparser.RawConfigParser() #read email data from config parser
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
        lender_list (list): list of lender,
        status_list (list): list of status,
        longest_term_buyrate_list (list): list of longest term buyrate,
        buy_rate_3_list (list): list of 3 month buyrate,
        buy_rate_6_list (list): list of 6 month buyrate,
        buy_rate_9_list (list): list of 9 month buyrate,
        buy_rate_12_list (list): list of 12 month buyrate,
        buy_rate_18_list (list): list of 18 month buyrate,
        highest_funding_amount_list (list): list of highest funding amount,
        longest_term_list (list): list of longest term,
        ad_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        doc_lis (list): list of funding stipulations
    """
    #subjects_stip = [msg.subject for msg in mailbox.fetch(AND(from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True))]
    subjects_temp = [msg.subject for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]
    sender_list_temp = [msg.from_ for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]
    date_list_temp = [msg.date_str for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]
    reciever_list_temp = [msg.to for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]
    html_content_temp=[msg.html for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(seen = False,from_='DoNotReply@forafinancial.com',subject='approval has been updated ',all=True),mark_seen=False)]

    subjects_temp = [' '.join(i.split()) for i in subjects_temp]
    html_content_temp = [' '.join(i.split()) for i in html_content_temp]
    print("The length of Fora Approval is ", len(subjects_temp))
    opportunity_name_list = [re.findall(r"(?<= \()(.*)(?=\) approval has )",op)[0] for op in subjects_temp]
   
    longest_term = [];highest_funding = [];buy_rate_3= [];buy_rate_6= [];buy_rate_9 = [];buy_rate_12 = []; buy_rate_18 = []
    longest_term_buy = [];ad_notes_list = [];doc_lis=[]

    for i in range(len(subjects_temp)):
        # html_data = html2text.html2text(html_content_temp[0])
        doc = re.findall(r'(?=We require)(.*)(?=Program Type:)',html_content_temp[0])
        #pre approved amount
        if doc ==[]:
           doc_lis.append(np.nan)
        else:
           for a_tuple in doc:
               doc_lis.append(doc[0])
        

        df = pd.read_html(html_content_temp[i],header=0)
        df = df[-1]
        df['Preapproved Amount'] = df['Preapproved Amount'].str.replace('$', '').str.replace(',','')
        high_list = list(df['Preapproved Amount'])
        
        high_list = [float(h) for h in high_list]
        high = max(high_list)
        highest_funding.append(high)
        
        term = max(df['Projected Turn'])
        buy = max(df['Buy Rate'])
        longest_term_buy.append(buy)
        longest_term.append(term)
        df.columns =[column.replace(" ", "_") for column in df.columns]
        
        df = df.drop_duplicates(subset=["Projected_Turn"])

        #buy rate extraction
        buy_3 =0;buy_6 =0;buy_9 =0;buy_12 =0;buy_18 =0;
        for x in range(len(df)):
            if df.Projected_Turn[x] == 3:
                buy_3 = df.Buy_Rate[x]
                buy_rate_3.append(buy_3)
            if df.Projected_Turn[x] == 6:
                buy_6 = df.Buy_Rate[x]
                buy_rate_6.append(buy_6)
                # print(buy_6)
            if df.Projected_Turn[x] == 9:
                buy_9 = df.Buy_Rate[x]
                buy_rate_9.append(buy_9)
                # print(buy_9)
            if df.Projected_Turn[x] == 12:
                buy_12 = df.Buy_Rate[x]
                buy_rate_12.append(buy_12)
            if df.Projected_Turn[x] == 18:
                buy_18 = df.Buy_Rate[x]
                buy_rate_18.append(buy_18)
                
        if buy_3 == 0:
            buy_rate_3.append(np.nan)
        if buy_6 == 0:
            buy_rate_6.append(np.nan)
        if buy_9 == 0:
            buy_rate_9.append(np.nan)
        if buy_12 ==0:
            buy_rate_12.append(np.nan)
        if buy_18 ==0:
            buy_rate_18.append(np.nan)    
            
 
        reciever_data = ""
        if type(reciever_list_temp[i]) == tuple or type(reciever_list_temp[i]) == list:
            for x in reciever_list_temp[i]:
                reciever_data = reciever_data + x+", "
            reciever_data = reciever_data[:-2]
        else:
            reciever_data = reciever_list_temp[i]

        html_data = html_content_temp[i]
        html_data = html_data.split(r'</style>', 1)[1]

        app_dec_notes = f"""<p>From: {sender_list_temp[i]}</p>
        <p>Date: {date_list_temp[i]}</p>
        <p>To: {reciever_data}</p>
        <p>Subjects: {subjects_temp[i]}</p>

        {html_data}
        """
        ad_notes_list.append(app_dec_notes)
        
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy2vNQAR',
        "status":'Approved',
        "highest funding amount":highest_funding,
        "longest term":longest_term,
        "longest term buyrate":longest_term_buy,
        "3 month buyrate":buy_rate_3,
        "6 month buyrate":buy_rate_6,
        "9 month buyrate":buy_rate_9,
        "12 month buyrate":buy_rate_12,
        "18 month buyrate":buy_rate_18,
        "approval decline notes":ad_notes_list,
        "subjects":subjects_temp,
        "subjects_uid":subjects_uid, 
        "Funding_Stipulations__c":doc_lis})
    
      
    df = df.drop_duplicates(subset=['opportunity',"highest funding amount","longest term"])
    df.to_csv("Fora_Approval.csv", index = False)
   
def main_fora_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        None
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    mail_content(inbox)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Fora_Approval Time in minutes: ",execution_time//60)

    
if __name__=="__main__":
    
    main_fora_approval()
    