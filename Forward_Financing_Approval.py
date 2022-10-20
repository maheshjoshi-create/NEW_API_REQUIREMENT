import configparser
import regex as re
# import html2text
import numpy as np
import pandas as pd
from imap_tools import MailBox, AND, OR
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup

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
        longest_term_list (list): list of longest term,
        application_decline_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        doc_list (list): list of Stipulations 
    """
    sen_lis=['newdecisions@forwardfinancing.com','pbunker@forwardfinancing.com','mjackson@forwardfinancing.com']
    sender_list = [msg.from_ for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)]
    mail_html=[msg.html for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)] 
    mail_text = [msg.text for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)] 
    n_subjects = [msg.subject for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in  mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='- Approved Offers',all=True),mark_seen=False)]
    high_funding_amt = [];long_term = [];ad_notes_list = [];doc_lis=[];opportunity_name_list=[];longest_term_buy=[];status=[]
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    
    # mail_html = [' '.join(i.split()) for i in mail_html] 
    mail_text = [' '.join(i.split()) for i in mail_text] 
    driver = webdriver.Chrome(ChromeDriverManager().install())
    print("The Length of Forward_Financing_Approval is",len(subjects_list))
    #  (?<=Closing Documents<\/font><\/u><\/i><\/div>)(.*?)([\s\S]*)(?=<div style="color: rgb\(0, 0, 0\); font-family: arial; font-size: 12pt;"><font face="Verdana, Helvetica, sans-serif" size="2"><i><\/i>)
    for i in range(len(subjects_list)):
        if('Re:'in subjects_list[i]):
            opportunity_name_list.append(re.findall(r"(?<=: )(.*)(?= - )",subjects_list[i])[0])
        else:
            opportunity_name_list.append(re.findall(r"(.*)(?= - )",subjects_list[i])[0])
        
        # html_data = html2text.html2text(mail_html[i])
        try:
            doc = re.findall(r'(?<=Closing Documents - )(.*)(?=Thank you)',mail_text[i])[0].replace(';','<br>')
        except:
            doc = re.findall(r'(?<=Closing Documents)(.*)(?=Thank you)',mail_text[0])[0].replace('>','').replace('*','').replace(',','<br>').replace('-','')
        # print(doc)   (?<=Closing Documents<\/u>)(.*)(?=<br><br><div style)
        # doc = re.findall(r'(?<=Closing Documents - )(.*)(?=Thank you)',mail_text[0])[0].replace(';','<br>')
        if doc ==[]:
            doc_lis.append(np.nan)
        else:
            doc_lis.append(doc)
        #     for a_tuple in doc:
        #         # print(i)
        #         doc_lis.append(a_tuple[1])
                # print(a_tuple[1])
        
        try:
            soft_offer = re.findall(r'(?<=Max Amount: \*\$)(.*)(?=Max Term)',mail_text[i])[0].replace('*','').replace('>','').replace(',','')
            high_funding_amt.append(soft_offer)
        except:
            soft_offer = re.findall(r'(?<=Max Amount: \$)(.*)(?=Max Term)',mail_text[i])[0].replace(',','')
            high_funding_amt.append(soft_offer)
        
        term = re.findall(r'(?<=Max Term:)(.*)(?=Program:)', mail_text[i])[0].replace('*','').replace('>','')
        
        # term = re.findall(r'(?<=Max Term:  \*\*)(.*)(?=\*\*)', html_data)[0]
        #          buy = re.findall(r'[1-9]+\.[0-9]+',html_data)[0]
        term = int(term)//20
        
        
        long_term.append(term) 
        match = re.findall(r'(?=https:\/\/bank.forwardfinancing.com\/offers\/)(.*)(?=Closing)', mail_text[i])
        driver.get(match[0])
        time.sleep(2)
        try:
            reset = driver.find_element_by_xpath('//*[@id="tab-content-by_approval"]/div[1]/form/div[1]/div[2]/a').send_keys(Keys.ENTER)
            # buy = driver.find_element_by_xpath('//*[@id="by_approval_total_factor_rate"]')
            # buy = driver.find_element_by_xpath('//*[@id="tab-content-by_approval"]/div[1]/form/div[4]/div/div[2]/div')
            # buy = driver.find_element_by_xpath('//*[@id="tab-content-by_approval"]/div[1]/form/div[4]/div/div[2]')
            buy = driver.find_element_by_class_name('static-variable')
            b = buy.find_element_by_xpath('//*[@id="by_approval_total_factor_rate"]')
            driver.execute_script("arguments[0].disabled=false;",b)
            bu = buy.find_element_by_xpath('//*[@id="by_approval_total_factor_rate"]').get_attribute("value")
            longest_term_buy.append(bu)
            status.append("Approved")
        except:
            longest_term_buy.append(np.nan)
            status.append("Expired Approval")
            
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
            
            
                    {mail_html[i]}
                
                    """
   
        app_dec_notes = re.sub(r"(?<=<style>)[\S\s\n]*(?=</style>)",'',app_dec_notes,1)
        ad_notes_list.append(app_dec_notes)
        
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002CyCvaQAF',
        "status":status,
        "highest funding amount":high_funding_amt,
        "longest term":long_term,
        "longest term buyrate":longest_term_buy,
        "3 month buyrate":np.nan,
        "6 month buyrate":np.nan,
        "9 month buyrate":np.nan,
        "12 month buyrate":np.nan,
        "18 month buyrate":np.nan,
        "approval decline notes":ad_notes_list,
        "Funding_Stipulations__c":doc_lis,
        "subjects":subjects_list,
        "subjects_uid":subjects_uid
       })
    df.to_csv("Forward_Financing_Approval.csv", index = False)


def main_forwad_financing_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        None
    """
    inbox=login_to_gmail()
    mail_content(inbox)
    

if __name__=="__main__":
    main_forwad_financing_approval()
