# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 09:44:46 2021

@author: Swara_Kaivu
"""

import configparser
import regex as re
import html2text
import math
import numpy as np
import pandas as pd
from imap_tools import MailBox, AND
def login_to_gmail():
    config=configparser.RawConfigParser()
    config.read('config.ini')
    
        # Login to Gmail
    mailbox = MailBox('imap.gmail.com')
    mailbox.login(config['GMAIL']['EMAIL'],config['GMAIL']['PASSWORD'], initial_folder='INBOX')
    return mailbox
def mail_content(mailbox):
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(subject='Congratulations! Your deal for'),AND(subject='has been approved External Id'),seen = False,from_='smartbizportal@everestbusinessfunding.com',all=True),mark_seen=False)]
    high_funding_amt = [];long_term = [];buy_rate=[];ad_notes_list = [];doc_lis=[]
    
    subjects_list = [' '.join(i.split()) for i in n_subjects]
    opportunity_name_list = [re.findall(r"(?<= for )(.*?)(?= has)",i)[0] for i in subjects_list]
    # html_data = html2text.html2text(mail_content[0])
    # print(html_data)
    print("The length of Everest Approval is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    for i in range(len(subjects_list)):
        d = ' '.join(mail_content[i].split())
        doc = re.findall(r'(?<=Please note any upsells should be noted in the INITIAL request for contracts. <\/p>)(.*?)([\s\S]*)(?= <p> <strong style="text-decoration: underline; color: red; font-weight: 700;">Note:<\/strong>)',d)
        if doc ==[]:
            doc_lis.append(np.nan)
        else:
            for a_tuple in doc:
                doc_lis.append(a_tuple[1])
        html_data = html2text.html2text(mail_content[i])
        #soft_offer = re.findall(r'(?:\d+\.)?\d+,\d+',html_data)[0]
        soft_offer = re.findall(r'(?<=Soft Offer: \*\*\$)(.*?)(?=\*\* )',html_data)[0]
        soft_offer =eval(soft_offer.replace(',',''))
        buy = re.findall(r'[1-9]+\.[0-9]+',html_data)[0]
        try:
            term = re.findall(r'(?<=Term Length: \*\*)(.*?)(?=\*)', html_data)[0]
            term = int(term)/21
        except:
            term = re.findall(r'(?<=Term Length: \*\*)(.*?)(?=Days)', html_data)[0]
            term = int(term)/21

        
        
        term = math.ceil(term)
        # print(term)
        high_funding_amt.append(soft_offer)
        long_term.append(term)
        buy_rate.append(buy)
        
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
        ad_notes_list.append(app_dec_notes)
    #print(len(opportunity_name_list),len(high_funding_amt),len(long_term),len(ad_notes_list),len(subjects_list),len(subjects_uid),len(doc_lis))    
    return opportunity_name_list,high_funding_amt,long_term,buy_rate,ad_notes_list,subjects_list, subjects_uid,doc_lis
def create_csv(o,h,term,buy,adn,subject,subjects_uid,doc_lis):
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002Cy2r3QAB',
        "status":"Approved",
        "highest funding amount":h,
        "longest term":term,
        "longest term buyrate":buy,
        "3 month buyrate":np.nan,
        "6 month buyrate":np.nan,
        "9 month buyrate":np.nan,
        "12 month buyrate":np.nan,
        "18 month buyrate":np.nan,
        "approval decline notes":adn,
        "Funding_Stipulations__c":doc_lis,
        "subjects":subject,
        "subjects_uid":subjects_uid
        })
    df.to_csv("Everest_Approval.csv", index = False)
def main_everest_approval():
    inbox=login_to_gmail()
    o,h,term,buy,adn,subjects_list, subjects_uid,doc_lis = mail_content(inbox)
    create_csv(o,h,term,buy,adn,subjects_list, subjects_uid,doc_lis)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_everest_approval()


