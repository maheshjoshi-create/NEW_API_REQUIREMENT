# -*- coding: utf-8 -*-
"""
Created on Mon Jan 24 10:02:06 2022

@author: Swara_Kaivu
"""

import pandas as pd 
import configparser
import regex as re
from imap_tools import MailBox, AND
import time
import  numpy as np
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys



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

def mail_content(mailbox):
    """ This function extracts the data from the email and returns the data in the form of list.
    
    Args:
        mailbox: MailBox object
        
    Returns:
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(subject='Full Approval - Deal Status Report -',seen = False, all=True),mark_seen=False)]
    
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    print("The Length of Business Backer Aprroval email is",len(subjects_list))


    opportunity_name_list = [re.findall(r"(?<=\d{6} - )(.*)(?= dba)",i)[0] for i in subjects_list]
    mail_content = [' '.join(i.split()) for i in mail_content]
    high_fund=[];long_term=[];ad_notes_list=[];status_lis=[];stips=[]
    driver = webdriver.Chrome(ChromeDriverManager().install())
   
    c_str = 'https://partners.businessbacker.com/'
    for i in range(len(subjects_list)):
        match = re.findall(r'(?<=https:\/\/partners.businessbacker.com\/)(.*)(?=" target="_blank">Configure Your Offer)', mail_content[i])
        if match != []:
            print(opportunity_name_list[i])
            match = c_str + match[0]
            # m.append(match)
            driver.get(match)
            time.sleep(8)
            # row_cl = driver.find_elements_by_class_name('row')
            
            
            
            try:
                slide_div = driver.find_elements_by_class_name('vue-slider')
                max_sli1 = slide_div[1].find_element_by_class_name('vue-slider-sr-only').get_attribute("max") 
                max_sli2 = slide_div[3].find_element_by_class_name('vue-slider-sr-only').get_attribute("min") 
                max_sli3 = slide_div[5].find_element_by_class_name('vue-slider-sr-only').get_attribute("max") 

                row_cl = driver.find_elements_by_class_name('row')
                # fir_slide = row_cl[13].find_element_by_xpath('//*[@id="quoteConfig"]/div[1]/div[2]/div[4]/div/div[6]/div[2]/input').send_keys(n_sli1)  # find the index of row_cl
                first_slide = row_cl[9].find_element_by_xpath('//*[@id="quoteConfig"]/div[1]/div[2]/div[4]/div/div[2]/div[2]/input').send_keys(max_sli1)    
                sec_slide = row_cl[11].find_element_by_xpath('//*[@id="quoteConfig"]/div[1]/div[2]/div[4]/div/div[4]/div[2]/input').send_keys(Keys.BACKSPACE)
                thi_slide = row_cl[13].find_element_by_xpath('//*[@id="quoteConfig"]/div[1]/div[2]/div[4]/div/div[6]/div[2]/input').send_keys(max_sli3)
                time.sleep(2)
                high = re.findall(r'(?<=Purchase Price \$ \$)(.*)', row_cl[15].text)[0].replace(',','')
                term = re.findall(r'(?<=Expected Term )(.*)(?= mo)', row_cl[15].text)[0].replace(',','')
                sti = row_cl[3].find_element_by_xpath('//*[@id="quoteConfig"]/div[1]/div[1]/div[3]/div').get_attribute('innerHTML')
                stip = re.findall(r'(?<=Paperwork).*', sti)
                
                # stip =  row_cl[3].text
                #buy = re.findall(r'(?<=Factor )(.*)', row_cl[15].text)[0].replace(',','')
                high_fund.append(high)
                long_term.append(term)
                stips.append(stip[0])
                #long_buyrate.append(buy)
                status_lis.append('Approved')
                
            except:
                high_fund.append(np.nan)
                long_term.append(np.nan)
                stips.append(np.nan)
                #long_buyrate.append(np.nan)
                status_lis.append('Expired Approval')

                pass
            # m.append(match)  
        else:
            print("No offer link present")
            high_fund.append(np.nan)
            long_term.append(np.nan)
            stips.append(np.nan)
            #long_buyrate.append(np.nan)
            status_lis.append('Approved')
            
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
        
    df = pd.DataFrame({"opportunity":opportunity_name_list,
    "lender":'0014W00002Cy2qsQAB',
    "status":status_lis,
    "highest funding amount":high_fund,
    "longest term":long_term,
    "longest term buyrate":np.nan,
    "3 month buyrate":np.nan,
    "6 month buyrate":np.nan,
    "9 month buyrate":np.nan,
    "12 month buyrate":np.nan,
    "18 month buyrate":np.nan,
    "approval decline notes":ad_notes_list,
    "Funding_Stipulations__c":stips,
    "subjects":subjects_list,
    "subjects_uid":subjects_uid
    })
    df.to_csv("Business_Backer_Approval.csv", index = False)
    return subjects_list,subjects_uid
            
def main_business_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        None
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    subjects_list,subjects_uid= mail_content(inbox) 
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Business_Approval Time in minutes: ",execution_time//60)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_business_approval()

     