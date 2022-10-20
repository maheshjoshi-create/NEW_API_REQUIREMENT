
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 21 10:56:07 2022

@author: user
"""

import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND,NOT
from bs4 import BeautifulSoup 
import time
import  numpy as np
import math
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

def login_to_gmail():
    """This function is used to login to the gmail account and return the mailbox object.
    
    Args:
        None
        
    Returns:
        mailbox: MailBox object
    """
    # Read the config.ini file
    config=configparser.RawConfigParser()
    config.read('config.ini')

    # Login to Gmail
    mailbox = MailBox('imap.gmail.com')
    mailbox.login(config['GMAIL']['EMAIL'],config['GMAIL']['PASSWORD'], initial_folder='INBOX')

    return mailbox

def mail_login(mailbox):
    """ This function extracts the data from the email and returns the data in the form of list.
    
    Args:
        mailbox: MailBox object
        
    Returns:
        n_subjects (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """

    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(NOT(subject='Out Of Office'),subject='Application Approved:'),AND(subject='- Toggle towards your deal'),seen = False, all=True),mark_seen=False)]
    n_subjects = [' '.join(i.split()) for i in n_subjects]
    mail_content1 = [' '.join(i.split()) for i in mail_content]
    high_funding_amt = [];long_term = [];longterm_buy_rate=[];ad_notes_list = [];buyrate_3=[];buyrate_6=[];buyrate_9=[];buyrate_12=[];buyrate_18=[];stips=[]

    s = 'https://pricecalc.pearlcapital.com/'
    opportunity_name_list = [re.findall(r"(?<=Approved: )(.*)(?= - )",i)[0] for i in n_subjects]
    print("_"*50)
    print("The Length of Pearl Funding Aprroval email is",len(n_subjects))
    for i in opportunity_name_list:
        print(i)
    print("_"*50)
    driver = webdriver.Chrome(ChromeDriverManager().install())
    for j in range(len(n_subjects)): 
        # print(j)
        reciever_data = ""
        if type(reciever_list[j]) == tuple or type(reciever_list[j]) == list:
            for x in reciever_list[j]:
                reciever_data = reciever_data + x+", "
            reciever_data = reciever_data[:-2]
        else:
            reciever_data = reciever_list[j]
            
        app_dec_notes = f"""<p>From: {sender_list[j]}</p>
        <p>Date: {date_list[j]}</p>
        <p>Subject: {n_subjects[j]}</p>
        <p>To: {reciever_data}</p>
    
    
        {mail_content1[j]}
        """
        # app_dec_notes = re.sub("",'',app_dec_notes)
        ad_notes_list.append(app_dec_notes)
        
        try:

            m = re.findall(r'(?<=https:\/\/pricecalc.pearlcapital.com\/)(.*)(?=" style)',mail_content[j])
            if '" data-auth=' in m[0]:
                m1 = re.findall(r'(.*)(?=" data-auth=)',m[0])
                match = s+m1[0]
            elif 'https://' in m[0]:
                match = s+m[1]           
            else:
                match = s+m[0]
            # print(match)
            driver.get(match)
            time.sleep(4)
            st = ''
            source  = BeautifulSoup(driver.page_source)
            time.sleep(4)
            
            
            all_div = source.find_all('div',{'id':'sliderContainer3'})
            new_all_div =[]
            for i in all_div:
                new_all_div.append(i.text)
                
            new_all_div = [' '.join(i.split()) for i in new_all_div]
            new_all_div = new_all_div[0]
            mon = re.findall(r'\d+ Months|\d+.\d+ Months',new_all_div)
            mon = [float(sub.split(' ')[0]) for sub in mon]
            buy = re.findall(r'(?<=Months\/)\d+.\d+(?=x)',new_all_div)
            lt = math.ceil(max(mon))
            lt_buy = max(buy) 
            long_term.append(lt)
            
            longterm_buy_rate.append(lt_buy)
            sam_df = pd.DataFrame({"months":mon,
                                  "buy":buy})
            three_months_buyrate=0
            six_months_buyrate=0
            nine_months_buyrate=0
            twelve_months_buyrate=0
            eighteen_months_buyrate=0
            
            st = str(source)
            sti = re.findall(r'(?<=data-content=")(.*)(?=data-html)', st)
            # sti = "<p>"+sti[0]+"</p>"
            sti = sti[0].replace('&lt;li&gt;','').replace("&lt;/li&gt;",'<br>').replace('"','')
            stips.append(sti)
            #print("buy rate:",lt_buy,"long term:",lt,"stipes:",sti)
            
            
            
            for i in range(0,len(sam_df)):
                
                if(sam_df.iloc[i]['months']== 3.0):
                    three_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_3.append(three_months_buyrate)
            
                elif (sam_df.iloc[i]['months']==6.0):
                    six_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_6.append(six_months_buyrate)
        
                elif (sam_df.iloc[i]['months']==9.0):
                    nine_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_9.append(nine_months_buyrate)
        
                elif (sam_df.iloc[i]['months']==12.0):
                    twelve_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_12.append(twelve_months_buyrate)
                    
                elif (sam_df.iloc[i]['months']==13.0):
                    eighteen_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_18.append(eighteen_months_buyrate)
        
                elif (sam_df.iloc[i]['months']==18.0):
                    eighteen_months_buyrate=sam_df.iloc[i]['buy']
                    buyrate_18.append(eighteen_months_buyrate)
                    
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
                
            
            h = re.findall(r'(?<=\$)\d+.\d+(?=Advance)', new_all_div)[0].replace(',','')
            high_funding_amt.append(h)
            
            if h == []:
                high_funding_amt.append(np.nan)
            
            
            
        except:
            stips.append(np.nan)
            high_funding_amt.append(np.nan)
            buyrate_18.append(np.nan)
            buyrate_12.append(np.nan)
            buyrate_9.append(np.nan)
            buyrate_6.append(np.nan)
            buyrate_3.append(np.nan)
            long_term.append(np.nan)
            longterm_buy_rate.append(np.nan)
            pass
    
        
    # driver.close()
        
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy2qfQAB',
        "status":"Approved",
        "highest funding amount":high_funding_amt,
        "longest term":long_term,
        "longest term buyrate":longterm_buy_rate,
        "3 month buyrate":buyrate_3,
        "6 month buyrate":buyrate_6,
        "9 month buyrate":buyrate_9,
        "12 month buyrate":buyrate_12,
        "18 month buyrate":buyrate_18,
        "approval decline notes":ad_notes_list,
        "Funding_Stipulations__c":stips,
        "subjects":n_subjects,
        "subjects_uid":subjects_uid
        })    
    df.to_csv("Pearl_Approval.csv", index = False)
    return n_subjects,subjects_uid

def main_pearl_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        n_subjects (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    n_subjects,subjects_uid = mail_login(inbox)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Pearl_Approval Time in minutes: ",execution_time//60)
    return n_subjects,subjects_uid
    
if __name__=="__main__":
    main_pearl_approval()
        
    