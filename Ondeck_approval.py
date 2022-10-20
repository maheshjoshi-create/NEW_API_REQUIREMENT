# -*- coding: utf-8 -*-
"""
Created on Mon Oct  3 13:03:50 2022

@author: mahesh.joshi
"""

import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND
import html2text
import numpy as np

def login_to_gmail():
    # Read the config.ini file
    config=configparser.RawConfigParser() #read email data from config parser
    config.read('config.ini')

    # Login to Gmail
    mailbox = MailBox('imap.gmail.com')
    mailbox.login(config['GMAIL']['EMAIL'],config['GMAIL']['PASSWORD'], initial_folder='INBOX')
    return mailbox
def mail_content(mailbox):
    # from_='offers@ondeckcapital.com',
    subjects_temp = [msg.subject for msg in mailbox.fetch(AND(seen = False,subject='- Approved Loan Offers, Application Number -',all=True),mark_seen=False)]
    sender_list_temp = [msg.from_ for msg in mailbox.fetch(AND(seen = False,subject='- Approved Loan Offers, Application Number - ',all=True),mark_seen=False)]
    date_list_temp = [msg.date_str for msg in mailbox.fetch(AND(seen = False,subject='- Approved Loan Offers, Application Number -',all=True),mark_seen=False)]
    reciever_list_temp = [msg.to for msg in mailbox.fetch(AND(seen = False,subject='- Approved Loan Offers, Application Number -',all=True),mark_seen=False)]
    html_=[msg.html for msg in mailbox.fetch(AND(seen = False,subject='- Approved Loan Offers, Application Number -',all=True),mark_seen=False)]  
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(seen = False,subject='. - Approved Loan Offers, Application Number -',all=True),mark_seen=False)]
    
    subjects_list = [' '.join(i.split()) for i in subjects_temp]
    html_content=[];high_fund=[];long_term=[];buyrate_3=[];buyrate_6=[];buyrate_9=[];buyrate_12=[];buyrate_18=[];longest_term_buyrate=[];ad_notes_list = [];stips=[]
    opportunity_name_list=[]
    
    for i in range(len(subjects_list)):
        if(':' in subjects_list[i]):
           
            a = re.findall(r"(?<=: )(.*?)(?=- Approved Loan Offers, Application Number)",subjects_list[i])[0]
            opportunity_name_list.append(a)
        elif ('#' in subjects_list[i]):
           
            a = re.findall(r"(.*?)(?=#)",subjects_list[i])[0]
            opportunity_name_list.append(a)

        else:
            
            a = re.findall(r"(.*?)(?=- Approved Loan Offers, Application Number)",subjects_list[i])[0]
            opportunity_name_list.append(a)

    print("_"*50)
    print("The length of OnDeck_Approval is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    print("_"*50)
    
    html_ = [' '.join(i.split()) for i in html_]
    sti = [re.findall(r"(?<=to activate their online account: )(.*?)(?=Your merchant will then be)",i) if ("they will need to provide the following four items to activate their online account:" in i) else re.findall(r"(?<=line of the email.<p> )(.*?)(?=<h3><u>3. )",i) for i in html_]
             
    for i in range(len(subjects_list)):
        print(i)
        reciever_data = ""
        if type(reciever_list_temp[i]) == tuple or type(reciever_list_temp[i]) == list:
            for x in reciever_list_temp[i]:
                reciever_data = reciever_data + x+", "
            reciever_data = reciever_data[:-2]
        else:
            reciever_data = reciever_list_temp[i]
       
        app_dec_notes = f"""<p>From: {sender_list_temp[i]}</p>
        <p>Date: {date_list_temp[i]}</p>
        <p>Subject: {subjects_list[i]}</p>
        <p>To: {reciever_data}</p>


        {html_[i]}
        """
        ad_notes_list.append(app_dec_notes) 
        try:
            df = pd.read_html(html_[i], header = 0)
            df=df[2]
            df.columns =[column.replace(" ", "_") for column in df.columns]
            try:
                Ammount = max(df['Max_Loan_Amount:'])
                high_fund.append(Ammount)
            except Exception as e:
                print(e)
                
            try:
                Term = max(df["Term:"])    
                long_term.append(Term)
            except Exception as e:
                print(e)
            
            if sti[i]==[]:
                stips.append(np.nan)
            else:
                stips.append(sti[i][0])
                
            if 'Credit_Limit:' not in df:
               
                long_term_buyrate=0
                three_months_buyrate=0
                six_months_buyrate=0
                nine_months_buyrate=0
                twelve_months_buyrate=0
                eighteen_months_buyrate=0
                long_term_buyrate=df['Max_Sell_Rate:'][0]
                longest_term_buyrate.append(long_term_buyrate)
                
                for i in range(0,len(df)):
                    
                    
                    if(df.iloc[i]['Term:']==3):
                        three_months_buyrate=df.iloc[i]['Max_Sell_Rate:']
                        buyrate_3.append(three_months_buyrate)
                        print("buyrate_3 :",three_months_buyrate)
            
                    elif (df.iloc[i]['Term:']==6):
                        six_months_buyrate=df.iloc[i]['Max_Sell_Rate:']
                        buyrate_6.append(six_months_buyrate)
                        print("buyrate_6 :",six_months_buyrate)
            
                    elif (df.iloc[i]['Term:']==9):
                        nine_months_buyrate=df.iloc[i]['Max_Sell_Rate:']
                        buyrate_9.append(nine_months_buyrate)
                        print("buyrate_9 :",nine_months_buyrate)
            
                    elif (df.iloc[i]['Term:']==12):
                        twelve_months_buyrate=df.iloc[i]['Max_Sell_Rate:']
                        buyrate_12.append(twelve_months_buyrate)
                        print("buyrate_12 :",twelve_months_buyrate)
            
                    elif (df.iloc[i]['Term:']==18):
                        eighteen_months_buyrate=df.iloc[i]['Max_Sell_Rate:']
                        buyrate_18.append(eighteen_months_buyrate)
                        print("buyrate_18 :",eighteen_months_buyrate)
                        
                if three_months_buyrate == 0:
                    buyrate_3.append(np.nan)
                if six_months_buyrate == 0:
                    buyrate_6.append(np.nan)
                if nine_months_buyrate == 0:
                    buyrate_9.append(np.nan)
                if twelve_months_buyrate ==0:
                    buyrate_12.append(np.nan)
                if eighteen_months_buyrate ==0:
                    buyrate_18.append(np.nan)
                if long_term_buyrate == 0:
                    longest_term_buyrate.append(np.nan)
              
            else:    
                buyrate_3.append(np.nan)
                buyrate_6.append(np.nan)
                buyrate_9.append(np.nan)
                buyrate_12.append(np.nan)
                buyrate_18.append(np.nan)
                longest_term_buyrate.append(np.nan)
        except:
                buyrate_3.append(np.nan)
                high_fund.append(np.nan)
                buyrate_6.append(np.nan)
                long_term.append(np.nan)
                buyrate_9.append(np.nan)
                buyrate_12.append(np.nan)
                buyrate_18.append(np.nan)
                longest_term_buyrate.append(np.nan)
                
                
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy2uxQAB',
        "status":"Approved",
        "highest funding amount":high_fund,
        "longest term":long_term,
        "longest term buyrate":longest_term_buyrate,
        "3 month buyrate":buyrate_3,
        "6 month buyrate":buyrate_6,
        "9 month buyrate":buyrate_9,
        "12 month buyrate":buyrate_12,
        "18 month buyrate":buyrate_18,
        "approval decline notes":ad_notes_list,
        "Funding_Stipulations__c":stips,
        "subjects":subjects_list,
        "subjects_uid":subjects_uid
        })
    df.to_csv("OnDeck_Approval.csv", index = False)
    return subjects_list, subjects_uid


def main_on_deck_approval():
    inbox=login_to_gmail()
    subjects_list, subjects_uid = mail_content(inbox) 
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_on_deck_approval()
            
                