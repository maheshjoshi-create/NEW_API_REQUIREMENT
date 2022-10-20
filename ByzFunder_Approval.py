import time
import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND, OR
import time
from webdriver_manager.chrome import ChromeDriverManager
import html2text
import numpy as np
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import InvalidSessionIdException



def login_to_gmail():
    """This function is used to login to the gmail account and return the mailbox object.
    
    Args:
        None
        
    Returns:
        mailbox: MailBox object
    """
    
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
        highest_funding_amount_list (list): list of highest funding amount,
        longest_term_list (list): list of longest term,
        application_decline_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        doc_lis (list): list of stipulation 
    """
    from_list = ['byzfunder@byzfunder.com','d.kulikova@byzfunder.com', 'ilyaf@byzfunder.com', 'v.fritz@byzfunder.com', 'uw@byzfunder.com', 'jmills@byzfunder.com','j.grossman@byzfunder.com','swapnil.chavan@aress.com']
    subjects_temp = [msg.subject for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)] 
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    mail_data=[msg.text for msg in mailbox.fetch(AND(OR(from_=from_list),seen = False,subject='Application Approved - ',all=True),mark_seen=False)]
    mail_data1 = [' '.join(i.split()) for i in mail_content]
    
    high_funding_amt = [];long_term = [];buy_rate=[];ad_notes_list = [];status_list=[];doc_lis=[]
    subjects_list = [' '.join(i.split()) for i in subjects_temp]
    print("The length of ByzFunder approvals is",len(subjects_list))
    # extract company name from subject list
    opportunity_name_list = [re.findall(r"(?<=Application Approved - ).*$",i)[0] for i in subjects_list]

    driver = webdriver.Chrome(ChromeDriverManager().install())
    newtext_list = []
    for i in range(len(subjects_list)): 
        # extract offer link 
        # try:
         # match = [re.findall(r"(?<=ByzFunder Offers\<)(.*)(?=\>)",i) if ("ByzFunder Offers" in i) else re.findall(r"(?<=Tandem Advance Offers\<)(.*)(?=\>)",i) for i in mail_content]
         
         try:
             match = [re.findall(r"(?<=href=)(.*)(?<=ByzFunder Offers)",mail_data1[i])][0][0].lstrip('\"\'').lstrip('\"').replace(">ByzFunder Offers",'').rstrip("\'").replace("amp;",'')
             driver.get(match) 
             #print(match[i][0])
             time.sleep(5)
             
         except:
             match = [re.findall(r"(?<=href=)(.*)(?<=Tandem Advance Offers)",mail_data1[i])][0][0].lstrip('\"\'').lstrip('\"').replace(">ByzFunder Offers",'').rstrip("\'").replace("amp;",'')
             driver.get(match) 
             #print(match[i][0])
             time.sleep(5)
             
        
         
        # pageSource = wait.until(EC.presence_of_element_located((By.XPATH, '//*')))
         pageSource = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*')))
 
         pageSource = pageSource.get_attribute("outerHTML")
         # convert page source into text format
         newtext = BeautifulSoup( pageSource, "lxml").text 
         # d = ' '.join(mail_content[i].split())
         doc = re.findall(r'(?=Outstanding Stips)(.*)(?=\*This)',mail_data1[i])
         if doc ==[]:
             doc_lis.append(np.nan)
         else:
            doc_lis.append(doc[0])
         
         newtext_list.append(newtext)
         # if offer link is expire then append nan values and status is offer link expired
         if newtext == 'Offer link expired':
             status_list.append('Expired Approval')
             high_funding_amt.append(np.nan)
             buy_rate.append(np.nan)
             long_term.append(np.nan)
             #mailbox.move(subjects_uid[i], 'ByzFunder Offered Expired')
             #print('Offer link expired')
             
             # otherwise extract data
         else:
             # extract funding amount
             #print('Offer not expired')
             advance = re.findall(r'(?<=Advance\$)(.*?)(?=Longest)', newtext)
             #advance = re.findall(r'(?<=ACH\$)(.*?)(?= Advance)', newtext)
             # 
             
             # print("advance before max: ",advance)
             try:
                 advance=max(advance)
                 advance=eval(advance.replace(',',''))
             except:
                 advance=np.nan
                 pass
                 
             # print(advance)
             # extract long term
             
             long = re.findall(r'(?<=Longest Turn)(.*?)(?=\s)', newtext)
             try:
                 long=max(long)
             except:
                 long=np.nan
                 pass
             
             
             # extract buy rate
             
             
             buy = re.findall(r'(?<=Payments)(.*?)(?=Buy Rate)', newtext)
             try:
                 buy=eval(max(buy))
             except:
                 buy=np.nan
                 pass
             
     
             
             high_funding_amt.append(advance)
             
             long_term.append(long)
             buy_rate.append(buy)
             status_list.append('Approved')
             # print('Data extraction')
             
        
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
            
    #     except:
    #         doc_lis.append(np.nan)
    #         status_list.append(np.nan)
    #         high_funding_amt.append(np.nan)
    #         buy_rate.append(np.nan)
    #         long_term.append(np.nan)
    #         ad_notes_list.append(np.nan)
            
    # df = pd.DataFrame({"opportunity":opportunity_name_list,
    #     "lender":'0014W00002Cy1sGQAR',
    #     "status":status_list,
    #     "highest funding amount":high_funding_amt,
    #     "longest term":long_term,
    #     "longest term buyrate":buy_rate,
    #     "3 month buyrate":np.nan,
    #     "6 month buyrate":np.nan,
    #     "9 month buyrate":np.nan,
    #     "12 month buyrate":np.nan,
    #     "18 month buyrate":np.nan,
    #     "approval decline notes":ad_notes_list,
    #     "Funding_Stipulations__c":doc_lis,
    #     "subjects":subjects_list,
    #     "subjects_uid":subjects_uid
        
    #     })
    # df = df.dropna(subset=['approval decline notes'])
    # df.to_csv("Byz_Funder_Approval.csv", index = False)
       
        
    return opportunity_name_list,status_list,high_funding_amt,long_term,buy_rate,ad_notes_list,subjects_list, subjects_uid,doc_lis
def create_csv(o,s,h,term,buy,adn,subject,subjects_uid,doc_lis):
    """ This function creates a csv file and writes the data in the csv file.
    
    Args:
        o (list): list of opportunity name,
        s (list): list of status,
        h (list): list of highest funding amount,
        term (list): list of longest term,
        buy (list): list of buy rate,
        adn (list): list of approval notes,
        subject (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
        doc_lis (list): list of stipulations
        
    Returns:
        None
    """
    
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002Cy1sGQAR',
        "status":s,
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
    df = df.dropna(subset=['approval decline notes'])
    df.to_csv("Byz_Funder_Approval.csv", index = False)
    for i in o:
        print(i)
    
def main_byzfunder_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        None
    """
    inbox=login_to_gmail()
    o,s,h,term,buy,adn,subjects_list, subjects_uid,doc_lis = mail_content(inbox)
    create_csv(o,s,h,term,buy,adn,subjects_list, subjects_uid,doc_lis)
    return subjects_list, subjects_uid
if __name__=="__main__":
     main_byzfunder_approval()
