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
        longest_term_buyrate_list (list): list of longest term buyrate,
        ad_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    # sub_=['APPROVAL OFFER: Mantis Funding - ','PRE-APPROVAL TERMS: Mantis Funding -']
    no_sub=['Automatic reply:']
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(NOT(OR(subject = no_sub)),subject='PRE-APPROVAL TERMS: Mantis Funding -'),seen = False, all=True),mark_seen=False)]

    
    high_funding_amt = [];long_term = [];ad_notes_list = [];longterm_buy_rate = []
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    
    opportunity_name_list = [re.findall(r"(?<=PRE-APPROVAL TERMS: Mantis Funding - )(.*)",i)[0] for i in subjects_list]
    print("_"*50)
    print("The Length of Mantis Aprroval email is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    print("_"*50)

    for i in range(len(subjects_list)):
        html_data = html2text.html2text(mail_content[i])
        html2 = ' '.join(html_data.split())
        try:
            funding_amount= re.findall(r'(?<= Financing \| \$)(.*)(?= Factor )',html2)[0].replace(',','')
            Ofpay= re.findall(r'(?<=Payments \| )(.*)(?= Origination)',html2)[0]
            t1 = int(Ofpay)/20
            t1=math.ceil(t1)                                                                                
            factor=re.findall(r'(?<=Factor \| )(.*)(?= Amount Payable )',html2)[0]
            f1 = float(factor)-(0.10)
            
            high_funding_amt.append(funding_amount)
            long_term.append(t1)
            longterm_buy_rate.append(f1)
        except:
            high_funding_amt.append(np.nan)
            long_term.append(np.nan)
            longterm_buy_rate.append(np.nan)
            
            
        
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

    return opportunity_name_list,high_funding_amt,long_term,longterm_buy_rate,ad_notes_list,subjects_list, subjects_uid

def create_csv(o,funding_amount,t1,f1,adn,subject,subjects_uid):
    """ This function creates a csv file and writes the data in the csv file.
    
    Args:
        o (list): list of opportunity name,
        funding_amount (list): list of highest funding amount,
        t1 (list): list of longest term,
        factor (list): list of longest term buyrate,
        adn (list): list of application decline notes,
        subject (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        
        
    Returns:
        None
    """
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002Cy2r2QAB',
        "status":"Approved",
        "highest funding amount":funding_amount,
        "longest term":t1,
        "longest term buyrate":f1,
        "3 month buyrate":np.nan,
        "6 month buyrate":np.nan,
        "9 month buyrate":np.nan,
        "12 month buyrate":np.nan,
        "18 month buyrate":np.nan,
        "approval decline notes":adn,
        "Funding_Stipulations__c":np.nan,
        "subjects":subject,
        "subjects_uid":subjects_uid
        })
    df.to_csv("Mantis_Approval.csv", index = False)
def main_mantis_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    o,funding_amount,t1,f1,adn,subjects_list,subjects_uid= mail_content(inbox)
    create_csv(o,funding_amount,t1,f1,adn, subjects_list, subjects_uid)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Mantis_Approval Time in minutes: ",execution_time//60)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_mantis_approval()

