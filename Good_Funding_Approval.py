


import pandas as pd 
import configparser
import regex as re

from imap_tools import MailBox, AND
import time
import  numpy as np
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

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
    sender_list = [msg.from_ for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    n_subjects = [msg.subject for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    date_list = [msg.date_str for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    reciever_list = [msg.to for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    mail_content=[msg.html for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    subjects_uid = [msg.uid for msg in mailbox.fetch(AND(subject='is Approved!',seen = False, all=True),mark_seen=False)]
    
    high_funding_amt = [];long_term = [];ad_notes_list = [];stips=[];buyrate_3=[]
    buyrate_6=[];buyrate_9=[];buyrate_12=[];buyrate_18=[];longest_term_buyrate=[]
    subjects_list = [' '.join(i.split()) for i in n_subjects] 
    mail  = [' '.join(i.split()) for i in  mail_content] 
    
    opportunity_name_list = [re.findall(r"(?<=: )(.*)(?= is Approved)",i)[0] if(':' in i) else re.findall(r"(.*)(?= is Approved)",i)[0] for i in subjects_list]

    # opportunity_name_list = [re.findall(r"(?<=: )(.*)(?= is Approved)",i)[0] if(':' in i) else re.findall(r"(.*)(?= is Approved!)",i)[0] for i in subjects_list]
    print("The Length of Good funding Aprroval email is",len(subjects_list))

    driver = webdriver.Chrome(ChromeDriverManager().install())
    #com = ['RE:','Re:']
    for i in range(len(subjects_list)):
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

        #try:
        if (":" not in subjects_list[i]):
            print('Yes')
            print(subjects_list[i])
            
            # high = re.findall(r'(?<=<\/b>\$)(.*)(?=<br \/>)',mail_content[i])[0].replace(',','')
            # tr = re.findall(r'(?<=Maximum Duration:<\/b> )(.*)(?=<br \/>)',mail_content[i])[0]
            try:
                # Updated
                stip = re.findall(r'(?<=Credit stips:)(.*)(?=Thank you)',mail[i])[0]
            except:
                # Old
                stip = re.findall(r'(?<=<\/span><\/span> )(.*)(?= <span style="font-size: 14px;">)',mail[i])[0]
            # high_funding_amt.append(high)
            # long_term.append(tr)
            stips.append(stip)
            
            
            match = re.findall(r'(?<=href=")(.*)(?=" target="_blank">)', mail_content[i])
            if match ==[]:
                print("Inside if")
                match = re.findall(r'(?<=href=")(.*)(?=" rel)',mail_content[i] )
            driver.get(match[0])
            time.sleep(12)
            # print(match[0])
            try:
                di1 = driver.find_element_by_xpath('//div[@class="slds-m-top_small slds-m-bottom_x-large"]')
    
                new = di1.get_attribute('innerHTML')
                df = pd.read_html(new,header=0)
                df=df[0]
                # print(df)
                df.columns =[column.replace(" ", "_") for column in df.columns]
                column = df["Buy_Rate"]
                hi = df["Purchase_Price"]
                tr = df["Duration_(Months)"]
                hi = [int(i.replace('$','').replace(',','') ) for i in hi]
                hi = pd.Series(hi)
                # print(column)
                max_buy=0
                high =0
                term = 0 
                max_buy = column.max()
                high = hi.max()
                term = tr.max()
                # print(max_buy)
                # high=high.replace('$','').replace(',','')
                
                longest_term_buyrate.append(max_buy)
                high_funding_amt.append(high)
                long_term.append(term)
                
                three_months_buyrate=0
                six_months_buyrate=0
                nine_months_buyrate=0
                twelve_months_buyrate=0
                eighteen_months_buyrate=0
                
                for i in range(0,len(df)):
                        
                    
                    if(df.iloc[i]['Duration_(Months)']==3):
                        three_months_buyrate=df.iloc[i]['Buy_Rate']
                        buyrate_3.append(three_months_buyrate)
            
                    elif (df.iloc[i]['Duration_(Months)']==6):
                        six_months_buyrate=df.iloc[i]['Buy_Rate']
                        buyrate_6.append(six_months_buyrate)
            
                    elif (df.iloc[i]['Duration_(Months)']==9):
                        nine_months_buyrate=df.iloc[i]['Buy_Rate']
                        buyrate_9.append(nine_months_buyrate)
            
                    elif (df.iloc[i]['Duration_(Months)']==12):
                        twelve_months_buyrate=df.iloc[i]['Buy_Rate']
                        buyrate_12.append(twelve_months_buyrate)
            
                    elif (df.iloc[i]['Duration_(Months)']==18):
                        eighteen_months_buyrate=df.iloc[i]['Buy_Rate']
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
                if max_buy == 0:
                    longest_term_buyrate.append(np.nan)
                
            except:
                high_funding_amt.append(np.nan)
                long_term.append(np.nan)
                # longest_term_buyrate.append(np.nan)
                # stips.append(np.nan)
                buyrate_3.append(np.nan)
                buyrate_6.append(np.nan)
                buyrate_9.append(np.nan)
                buyrate_12.append(np.nan)
                buyrate_18.append(np.nan)
                longest_term_buyrate.append(np.nan)
                pass
            
            
 
        else:
            print('No')
            print(subjects_list[i])
          
            high_funding_amt.append(np.nan)
            long_term.append(np.nan)
            longest_term_buyrate.append(np.nan)
            stips.append(np.nan)
            buyrate_3.append(np.nan)
            buyrate_6.append(np.nan)
            buyrate_9.append(np.nan)
            buyrate_12.append(np.nan)
            buyrate_18.append(np.nan)
   
       
    driver.close()
    df = pd.DataFrame({"opportunity":opportunity_name_list,
        "lender":'0014W00002Cy1xEQAR',
        "status":"Approved",
        "highest funding amount":high_funding_amt,
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
    df.to_csv("Good_Funding_Approval.csv", index = False)

    return subjects_list,subjects_uid    

   
def main_good_funding_approval():
    """ This function is used to extract the data from the email.
    
    Args:
        None
    
    Returns:
        subjects_list: list of subjects,
        subjects_uid: list of subjects uid
    """
    inbox=login_to_gmail()
    start_Time=time.time()
    subjects_list, subjects_uid = mail_content(inbox)
    end_Time=time.time() # At the end of script
    execution_time=(end_Time-start_Time)
    print("Good_Funding_Approval Time in minutes: ",execution_time//60)
    
    return subjects_list, subjects_uid
if __name__=="__main__":
    main_good_funding_approval()
    
   
    
        
        
   
    

