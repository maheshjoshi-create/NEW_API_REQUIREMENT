import configparser
import regex as re
import html2text
import numpy as np
import pandas as pd
from imap_tools import MailBox, AND,OR
import time

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
        high_funding_amt (list): list of highest funding amount,
        long_term (list): list of longest term,
        buy_rate (list): list of buy rate,
        ad_notes_list (list): list of application decline notes,
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
        doc_lis (list): list of funding stipulations
    """
    sen_lis = ['stips@libertasfunding.com','stips@nationalbizcap.com']
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    mail_text=[msg.text for msg in mailbox.fetch(AND(OR(from_=sen_lis),seen = False,subject='Pre-Approval for',all=True),mark_seen=False)]
    

    high_funding_amt = [];long_term = [];buy_rate=[];ad_notes_list = []
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    mail_text = [' '.join(i.split()) for i in mail_text]
    # mail_content = [' '.join(i.split()) for i in mail_content]
    ht = [];doc_lis=[]
    opportunity_name_list = [re.findall(r"(?<= for )(.*)",i)[0] for i in subjects_list]
    print("The Length of Libertas Aprroval email is",len(subjects_list))
    for i in opportunity_name_list:
        print(i)
    for i in range(len(subjects_list)):
        
        html_data = html2text.html2text(mail_content[i])
        ht.append(html_data)
        
        doc = re.findall(r'(?<=<!--STIPS-->)(.*?)([\s\S]*)(?=<br><hr>)',mail_content[i])
        if doc ==[]:
            doc_lis.append(np.nan)
        else:
            for a_tuple in doc:
                # print(i)
                doc_lis.append(a_tuple[1])
        try:
            # soft_offer = re.findall(r'(?:\d+\.)?\d+,\d+.\d\d',html_data)[0]
            # soft_offer = re.findall(r'(?=\$\d+)(.+)(?= About Offer: )',mail_text[i])[0].split()
            # soft_offer = eval(soft_offer[0].replace(",",'').replace('$',''))
            df = pd.read_html(mail_content[i],match = 'Approved Amount')
            df2 = df[-1]
            amount = df2['Approved Amount'][0].replace('$','').replace(',','')
            high_funding_amt.append(amount)
        except:
            high_funding_amt.append(np.nan)
            
        
        try:            
            term = re.findall(r' \d{1} | \d{2} ', html_data)[0]
            long_term.append(term)
            br=float(re.findall(r'[1-9]+\.[0-9]+',html_data)[0])
            fr = df2['Points Built In'].iloc[0]
            fr=fr/100
            b1 =br-fr
            buy_rate.append(b1)
            
           
        except:
           
            long_term.append(np.nan)
            buy_rate.append(np.nan)
            
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
    return opportunity_name_list,high_funding_amt,long_term,buy_rate,ad_notes_list,subjects_list, subjects_uid,doc_lis

def create_csv(o,h,term,b1,adn,subject,subjects_uid,doc_lis):
    """ This function creates a csv file and writes the data in the csv file.
    
    Args:
        o (list): list of opportunity name,
        h (list): list of highest funding amount,
        term (list): list of longest term,
        adn (list): list of application decline notes,
        subject (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail,
        doc_lis (list): list of Stipulations
        
    Returns:
        None
    """
    df = pd.DataFrame({"opportunity":o,
        "lender":'0014W00002CxvPjQAJ',
        "status":"Approved",
        "highest funding amount":h,
        "longest term":term,
        "longest term buyrate":b1,
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
    df.to_csv("Libertas_Approval.csv", index = False)
def main_libertas_approval():
    
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        None
    """
    inbox=login_to_gmail()
    start_Time=time.time()
    o,h,term,b1,adn,subjects_list, subjects_uid,doc_lis = mail_content(inbox)
    create_csv(o,h,term,b1,adn,subjects_list, subjects_uid,doc_lis)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print('Libertas Approval is completed')
    print("Libertas_Approval Time in minutes: ",execution_time//60)
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_libertas_approval()
