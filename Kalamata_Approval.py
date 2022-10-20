
import configparser
import regex as re
import pandas as pd
from imap_tools import MailBox, AND
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
        subject_list (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    
    # list_=['samantha.singh@kalamatacapitalgroup.com','ryannv@national.biz','ryan.biango@kalamatacapitalgroup.com','pasquale.mastroviti@kalamatacapitalgroup.com','updates@kalamatacapitalgroup.com','msteam@kalamatacapitalgroup.com','carl.reed@kalamatacapitalgroup.com','stephs@national.biz','david.aiosa@kalamatacapitalgroup.com','mark.schrews@kalamatacapitalgroup.com']
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(AND(subject='Application Approved: '),AND(subject=' / '),seen = False,all=True),mark_seen=False)]
     
   
    high_funding_amt = [];long_term = [];ad_notes_list = [];opportunity_name_list=[]; longterm_buy_rate =[]
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    subjects_list = [i.replace('\u200b','') for i in subjects_list] #we have replace '\u200b' with empty space as it creates a problem while moving the mail
    
    print("_"*50)
    print("The length of Kalamata Approval is ", len(subjects_list))
   
    for i in range(len(subjects_list)):

        if "KCG" in mail_content[i]:
            # print(i)
            # print(subjects_list[i])
            
            try: 
                df = pd.read_html(mail_content[i], match='Approval Amount')
                df_new = df[0]
                df=df[-1]
                amt = df[1].iloc[0].replace('$','').replace(',','')
                high_funding_amt.append(amt)
                opp = re.findall(r"(?<=Application Approved: )(.*)(?= \/ )",subjects_list[i])[0] 
                opportunity_name_list.append(opp)
                
                
                comission = df_new.Commission[df_new.Commission == "$0.00"].dropna().index.tolist()[0]
                buy_rate = df_new['Factor Rate']
                buy_rate = buy_rate.iloc[comission][-1]
                longterm_buy_rate.append(buy_rate)
                
                    
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
                
                for i in range(0,len(df)): 
                    
                    if(df.iloc[i][1]=='Daily'):
                        term=int(df.iloc[i-1][1])/20
                        long_term.append(term)
                        
                    if(df.iloc[i][1]=='Weekly'):
                        term=int(df.iloc[i-1][1])/4
                        long_term.append(term)
                    
                    if(df.iloc[i][1]=='Monthly'):
                        term=int(df.iloc[i-1][1])
                        long_term.append(term)
                     
                    
                
            except Exception as e:
                print('Error',e)
                long_term.append(np.nan)     
                high_funding_amt.append(np.nan)
                ad_notes_list.append(np.nan)
                opportunity_name_list.append(np.nan)
                longterm_buy_rate.append(np.nan)

        else:
            long_term.append(np.nan)     
            high_funding_amt.append(np.nan)
            ad_notes_list.append(np.nan)
            opportunity_name_list.append(np.nan)
            longterm_buy_rate.append(np.nan)

    
    for i in opportunity_name_list:     
        if type(i) == str:
            print(i)
        
    print("_"*50)        
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy0aJQAR',
        "status":"Approved",
        "highest funding amount":high_funding_amt,
        "longest term":long_term,
        "longest term buyrate":longterm_buy_rate,
        "3 month buyrate":np.nan,
        "6 month buyrate":np.nan,
        "9 month buyrate":np.nan,
        "12 month buyrate":np.nan,
        "18 month buyrate":np.nan,
        "approval decline notes":ad_notes_list,
        "Funding_Stipulations__c":np.nan,
        "subjects":subjects_list,
        "subjects_uid":subjects_uid
        })
    df = df.dropna(subset=['approval decline notes'])
    df.to_csv("Kalamata_Approval.csv", index = False)
    
    # if len(df)>0:
    #     is_exists = mailbox.folder.exists('Processed')
    #     if is_exists != True:
    #         mailbox.folder.create('Processed')
    #     for i in range(len(df)):
    #         print(df.subjects.iloc[i])
    #         print(df.subjects_uid.iloc[i])
    #         mailbox.move(str(df.subjects_uid.iloc[i]), 'Processed')
    return subjects_list,subjects_uid
           
def main_kalamata_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        subjects_list: List of subjects of the emails.
        subjects_uid: List of UIDs of the emails.
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    subjects_list,subjects_uid = mail_content(inbox)
    
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("kalamata_Approval Time in minutes: ",execution_time//60)
    return subjects_list,subjects_uid
    
if __name__=="__main__":
    main_kalamata_approval()

