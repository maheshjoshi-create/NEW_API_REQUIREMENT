import configparser
import regex as re
import pandas as pd
import html2text
from imap_tools import MailBox, AND,OR,NOT
import math
import time
import  numpy as np


def login_to_gmail():
    """This function is used to login to the gmail account and return the mailbox object.
    
    Args:
        None
        
    Returns:
        mailbox: MailBox object
    """
    config=configparser.RawConfigParser()
    config.read('config.ini')

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
        longest_term_buyrate_list (list): list of longest term buyrate,
        ad_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    # list_=['noreply@bittyadvance.com']  OR(from_=list_),
    no_sub=['Duplicate','has been approved','Congratulations!','REVISED.','COMPETING - ']  # SHOPPING -- Urgent
    
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='APPROVED - '),AND(subject=' ID:'),seen = False, all=True),mark_seen=False)]
    
    high_funding_amt = [];long_term = [];longterm_buy_rate=[];ad_notes_list = [];  opportunity_name_list=[]; funding_stipulations=[]
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    mail_content = [' '.join(i.split()) for i in mail_content] 

    for i in subjects_list:
        if 'Pending Submission'in i:
            opportunity_name_list.append(re.findall(r"(?<=-)(.*)",i)[0])
        else:
            opportunity_name_list.append(re.findall(r"(?<=APPROVED - )(.*)(?= ID:)",i)[0])
            
    # opportunity_name_list = [re.findall(r" (?<=-)(.*)",i) if 'Pending Submission' in i else re.findall(r"(?<=-)(.*)(?=ID)",i)  for i in subjects_list]
    print("_"*50)
    print("The Length of Bitty Aprroval email is",len(subjects_list))
    
    print("_"*50)
    
    for i in range(len(subjects_list)):
        html_data = html2text.html2text(mail_content[i])
        html2 = ' '.join(html_data.split())
       
        try:
            offer = re.findall(r'(?<=--- Offer ## \$ )(.*)(?= \| Commission ## )',html2)[0].replace(',','') 
            factor= re.findall(r'(?<=Factor ## )(.*)(?= \| Commission %)',html2)[0]
            commission= re.findall(r'(?<=Commission % ## )(.*)(?=.00 % Payback)',html2)[0]
            c1= int(commission)/100
            f1= float(factor)-(c1)
            term=re.findall(r'(?<=Term ## )(.*)(?= \| If )',html2)[0]
            t1 = int(term)/20
            t1=math.ceil(t1)
            try:
              stips=re.findall(r'(?<=account.<\/div>)(.*)(?=National Biz LLC)',mail_content[i])[0]
              funding_stipulations.append(stips)
            except:
                funding_stipulations.append(np.nan)
            
            high_funding_amt.append(offer)
            longterm_buy_rate.append(f1)
            long_term.append(t1)
            
        except:
            high_funding_amt.append(np.nan)
            longterm_buy_rate.append(np.nan)
            long_term.append(np.nan)
            funding_stipulations.append(np.nan)
            
            
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
    
    return opportunity_name_list,high_funding_amt,long_term,longterm_buy_rate,ad_notes_list,funding_stipulations,subjects_list, subjects_uid

def create_csv(o,offer,t1,f1,adn,stips,subject,subjects_uid):
    """ This function creates a csv file and writes the data in the csv file.
    
    Args:
        o (list): list of opportunity name,
        offer (list): list of highest funding amount,
        t1 (list): list of longest term,
        factor (list): list of longest term buyrate,
        adn (list): list of application decline notes,
        subject (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
        
    Returns:
        None
    """
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002Cy200QAB',
        "status":"Approved",
        "highest funding amount":offer,
        "longest term":t1,
        "longest term buyrate":f1,
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
    df.to_csv("Bitty_Approval.csv", index = False)
def main_bitty_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        subjects_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    o,offer,t1,f1,adn,stips,subjects_list,subjects_uid= mail_content(inbox)
    create_csv(o,offer,t1,f1,adn,stips,subjects_list, subjects_uid)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Bitty_Approval Time in minutes: ",execution_time//60)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_bitty_approval()
