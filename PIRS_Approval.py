import configparser
import regex as re
import numpy as np
import pandas as pd
from imap_tools import MailBox, AND
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
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
        n_subjects (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
   
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(from_='noreply@pirscapital.com',seen = False,subject='has been approved!',all=True),mark_seen=False)]
    high_funding_amt = [];long_term = [];ad_notes_list = [];longterm_buy_rate=[];buyrate_3=[];buyrate_6=[];buyrate_9=[];buyrate_12=[];buyrate_18=[]
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    mail_content = [' '.join(i.split()) for i in mail_content] 
    stips = [re.findall(r'(?<=Provide with signed contracts:<\/strong><\/p>)(.*)(?=<strong>PIRS Capital will take care of:)',i)[0] for i in mail_content]
    opportunity_name_list = [re.findall(r"(?<=\")(.*)(?=\" has been approved!)",i)[0] for i in subjects_list]
    print("_"*50)
    print("The Length of PIRS Aprroval email is",len(n_subjects))
    for i in opportunity_name_list:
        print(i)
    print("_"*50)
    s='https://pirsportal.com/public/calculator/'
    driver = webdriver.Chrome(ChromeDriverManager().install())
    
    for i in range(len(mail_content)):
        reciever_data = ""
        if type(reciever_list[i]) == tuple or type(reciever_list[i]) == list:
            for x in reciever_list[i]:
                reciever_data = reciever_data + x+", "
            reciever_data = reciever_data[:-2]
        else:
            reciever_data = reciever_list[i]
            
        app_dec_notes = f"""<p>From: {sender_list[i]}</p>
        <p>Date: {date_list[i]}</p>
        <p>Subject: {n_subjects[i]}</p>
        <p>To: {reciever_data}</p>


        {mail_content[i]}
        """
        app_dec_notes = re.sub(r'(?=An offer is provided)(.*)(?<=hide: all; })','',app_dec_notes)
        ad_notes_list.append(app_dec_notes) 

        m=re.findall(r'(?<=https:\/\/pirsportal.com\/public\/calculator\/)(.*)(?=">Calculate Your Approval)',mail_content[i])
        m=re.findall(r'(.*)(?=" class=)',m[0])[0]
        match=s+m
        driver.get(match)
        time.sleep(5)
        source  = BeautifulSoup(driver.page_source,features="lxml")
        all_div = source.find_all('div',class_='calculation_info_boxes_wrapper')
        new_all_div =[]
        for i in all_div:
                new_all_div.append(i.text)
        new_all_div = [' '.join(i.split()) for i in new_all_div]
        new_all_div = new_all_div[0] 
        buy = re.findall(r'(Buy Rate: \d+.\d+)',new_all_div)
        buy = [float(sub.split(' ')[2]) for sub in buy]
        mon = re.findall(r'(\d+ Month)',new_all_div)
        mon = [int(sub.split(' ')[0]) for sub in mon]
        lt = max(mon)
        lt_buy = max(buy)        
        long_term.append(lt)
        longterm_buy_rate.append(lt_buy)
        sam_df = pd.DataFrame({"months":mon,
                                  "buy":buy})
        # sam_df.duplicated(keep = 'last')
        sam_df.drop_duplicates(subset="months",inplace=True)
        sam_df.reset_index(inplace = True, drop = True)
        
        three_months_buyrate=0
        six_months_buyrate=0
        nine_months_buyrate=0
        twelve_months_buyrate=0
        eighteen_months_buyrate=0
        
        
        
        for j in range(0,len(sam_df)):
                
            if(sam_df.iloc[j]['months']== 3):
                three_months_buyrate=sam_df.iloc[j]['buy']
                buyrate_3.append(three_months_buyrate)
        
            elif (sam_df.iloc[j]['months']==6):
                six_months_buyrate=sam_df.iloc[j]['buy']
                buyrate_6.append(six_months_buyrate)
    
            elif (sam_df.iloc[j]['months']==9):
                nine_months_buyrate=sam_df.iloc[j]['buy']
                buyrate_9.append(nine_months_buyrate)
    
            elif (sam_df.iloc[j]['months']==12):
                twelve_months_buyrate=sam_df.iloc[j]['buy']
                buyrate_12.append(twelve_months_buyrate)
    
            elif (sam_df.iloc[j]['months']==18):
                eighteen_months_buyrate=sam_df.iloc[j]['buy']
                buyrate_18.append(eighteen_months_buyrate)
                
        if three_months_buyrate == 0:
            buyrate_3.append(np.nan)
            print("buyrate_3: ",len(buyrate_3))
        if six_months_buyrate == 0:
            buyrate_6.append(np.nan)
            print("buyrate_6: ",len(buyrate_6))
        if nine_months_buyrate == 0:
            buyrate_9.append(np.nan)
            print("buyrate_9: ",len(buyrate_9))
        if twelve_months_buyrate ==0:
            buyrate_12.append(np.nan)
            print("buyrate_12: ",len(buyrate_12))
        if eighteen_months_buyrate ==0:
            buyrate_18.append(np.nan)
            print("buyrate_12: ",len(buyrate_12))
            
        hi = re.findall(r'(?<=Advance: \$)\d+,\d+', new_all_div)[0].replace(',','')
        high_funding_amt.append(hi)
          
        if hi == []:
            high_funding_amt.append(np.nan)
    
      
    driver.close()
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy2wkQAB',
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
    df.to_csv("PIRS_Approval.csv", index = False) 
    return n_subjects,subjects_uid
    
def main_pirs_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        n_subjects (list): list of subject of mail,
        subjects_uid (list): list of uid of the mail
    """
    start_Time=time.time()
    inbox=login_to_gmail()
    n_subjects,subjects_uid = mail_content(inbox)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("PIRS_Approval Time in minutes: ",execution_time//60)
    return n_subjects,subjects_uid
    
if __name__=="__main__":
    main_pirs_approval()
        
