# General Libraries
import os
import win32com.client as win32
import pandas as pd

# Time Libraries
from datetime import datetime, timedelta
import time

# AI Libraries
from dotenv import load_dotenv
import google.generativeai as genai

# Environment Variables
load_dotenv()
import gspread
from google.oauth2.service_account import Credentials
SERVICE_ACCOUNT_FILE = "service_account.json" # service account file for subscription checking (not the optimal way to do authorization)
SHEET_NAME = "client_list" # sheet containing authorized users and subscription details

# Google Gemini AI Configuration
api_key = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-2.5-flash') # change between pro and flash when needed

# Check user authorization for JTmail
def check_license():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)

    sheet = client.open(SHEET_NAME).sheet1  
    records = sheet.get_all_records() 
    for row in records:
        if row.get("email", "").strip().lower() == os.getenv("EMAIL_ADDRESS"):
            if str(row.get("active", "")).strip().lower() == "true":
                return True
            else:
                return False
    return False

# Reads a CSV file and return its data
def read_spreadsheet():
    while True:
        file = input("Enter the file path to your spreadsheet (must be .csv): ").strip()
        file = file.strip('"').strip("'")
        file = os.path.normpath(file)
        if not os.path.exists(file):
            print(f"Error: File not found - {file}")
        else:
            data = pd.read_csv(file)
            return data

# Checks if an email has already been sent to the lead
def already_sent(email):
    if os.path.exists('sent.csv'):
        sent_data = pd.read_csv('sent.csv')
        return email in sent_data['Email'].values
    return False

# Checks if a company is in the Do-Not-Contact list
def do_not_contact(company):
    if os.path.exists('do_not_contact.txt'):
        with open('do_not_contact.txt', 'r', encoding='utf-8') as f:
            already_closed = [line.strip().lower() for line in f if line.strip()]
            for line in already_closed:
                if company.lower() in line:
                    return True
    return False

# Sends emails using Outlook
def send_email(data, sent_today, sent_total, email_type):
    mail_amount = int(input("Enter the number of emails you want to send: "))
    print()
    print("Select the industry being sent to: ") # Can be tailord to other desired industries as needed
    print("[1] Food & Beverage")
    print("[2] Sports Teams & Organizations")
    industry = str(input(">>> ")).strip()
    print()

    if email_type == "followup":
        print("Enter the minimun amount of days since last contact: ")
        days = int(input(">>> "))
        print()

    mail_sent = 0
    outlook = win32.DispatchEx('Outlook.Application')
    sent_bool = False

    for _, row in data.iterrows():
        sent_emails = []
        if mail_sent >= mail_amount:
            print(str(mail_sent) + "/" + str(mail_amount) + " emails sent successfully.")
            break
        email = str(row['Email'])
        if already_sent(email) and email_type == "cold":
            print(f"Skipping already contacted lead.")
            continue
        if email == "nan" or email == "" or pd.isna(email):
            print(f"Skipping lead with no email.")
            continue
        
        # Extract relevant information from the data
        first_name = str(row['First Name'])
        company = str(row['Company'])

        if do_not_contact(company):
            print(f"Skipping lead: {company} in the Do-Not-Contact list.")
            continue

        if email_type == "followup":
            icebreaker = ""
            today = datetime.now()
            raw_date = str(row['Date']).strip()
            sent_date = datetime.strptime(raw_date, "%m/%d/%Y %H:%M")
            diff = today - sent_date

            if (datetime.now() - datetime.strptime(str(row['Date']).strip(), "%m/%d/%Y %H:%M")).days >= days:
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = "Hi " + first_name + ","
                mail.HTMLBody = generate_email(first_name, icebreaker, email_type, industry)
                image_path = os.path.join(os.getcwd(), "YourPhotoOrLogo.png")
                image = mail.Attachments.Add(image_path)
                image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "profileimage")
                product_list = os.path.join(os.getcwd(), "YourProductList.pdf")
                mail.Attachments.Add(product_list)

                #Set the sender email
                for account in outlook.Session.Accounts:
                    if account.SmtpAddress.lower() == "youremail@provider.com":
                        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                        break

                mail.Send()  # Use mail.Display() to preview before sending or mail.Send() to send immediately
                print("Follow Up sent to " + email)
                sent_df = pd.read_csv('sent.csv')
                sent_df.loc[sent_df['Email'] == email, 'Follow Up'] = datetime.now().strftime("%m/%d/%Y %H:%M")
                sent_df.to_csv('sent.csv', index=False)
                sent_bool = True
                print("30 second cooldown, please wait...") # Cooldown to avoid spam filters
                time.sleep(30)
            else: continue

        else:
            company_desc = str(row['Company Description'])
            company_seo = str(row['Company SEO Description'])
            icebreaker = generate_icebreaker(company, company_desc, company_seo)
            print(f"Icebreaker: {icebreaker}")

            #Create and send the email
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "Hi " + first_name + ","
            mail.HTMLBody = generate_email(first_name, icebreaker, email_type, industry)
            image_path = os.path.join(os.getcwd(), "YourPhotoOrLogo.png")
            image = mail.Attachments.Add(image_path)
            image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "profileimage")
            product_list = os.path.join(os.getcwd(), "YourProductList.pdf")
            mail.Attachments.Add(product_list)

            # Set the sender email
            for account in outlook.Session.Accounts:
                if account.SmtpAddress.lower() == "youremail@provider.com":
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    break

            try: # Log the sent email and update apppropriate variables
                match industry:
                    case "1":
                        mail.Send()  # Use mail.Display() to preview before sending or mail.Send() to send immediately
                        print("Email sent to " + email)
                        sent_emails.append({
                            'First Name': first_name,
                            'Email': email,
                            'Company': company,
                            'Industry': "Food & Beverage",
                            'Date': datetime.now().strftime("%m/%d/%Y %H:%M"),
                            'Follow Up': "Pending"
                        })
                        new_df = pd.DataFrame(sent_emails)
                        new_df.to_csv('sent.csv', mode='a', header=False, index=False)
                    case "2":
                        mail.Send()  # Use mail.Display() to preview before sending or mail.Send() to send immediately
                        print("Email sent to " + email)
                        sent_emails.append({
                            'First Name': first_name,
                            'Email': email,
                            'Company': company,
                            'Industry': "Sports Team Organization",
                            'Date': datetime.now().strftime("%m/%d/%Y %H:%M"),
                            'Follow Up': "Pending"
                        })
                        new_df = pd.DataFrame(sent_emails)
                        new_df.to_csv('sent.csv', mode='a', header=False, index=False)
                    case _:
                        print("Industry not recognized")
            except Exception as e:
                print(f"Failed to send email to " + email + f": {e}")
            print("30 second cooldown, please wait...") # Cooldown to avoid spam filters
            time.sleep(30)
        mail_sent += 1

    if sent_bool:
        if mail_sent < mail_amount:
            print(str(mail_sent) + "/" + str(mail_amount) + " emails sent successfully. (Spreadsheet exhausted)")
        sent_today += mail_sent
        sent_total += mail_sent
        with open("sent_data.txt", "w") as file:
            file.write(str(sent_today) + ", " + str(sent_total) + ", " + datetime.now().strftime("%m/%d/%Y %H:%M"))

    else:
        print("0 emails sent successfully. (Spreadsheet exhausted)")
    return sent_today, sent_total

# Generates a personalized icebreaker using AI
def generate_icebreaker(company_name, company_desc, company_seo):
    prompt = f"""
    You are a friendly and professional email marketer writing a very short icebreaker for a cold email.

    Context:
    You discovered the company below and want to start a conversation.
    You work for a family-owned business that sells organic products.

    Details:
    - Company Name: {company_name}
    - Company Description: {company_desc}
    - Company SEO Description: {company_seo}

    Formatting Rules (mandatory):
    1. Company name: Convert to Title Case and remove suffix words like Inc., LLC, etc. When mentioning the company name, make it sound natural and conversational as if you were saying it out loud. Never remove the first word of the company name.
    2. Company Product: Extract the company's product from the company description and SEO description, mention it using less than five words, use 'and' instead of '+' or '&', and convert it to all lowercase (even if it is a proper noun like acai).
    3. If no clear product is found, skip the product phrase and use the fallback pattern.
    4. If the company name does not match the company description, use the "Company Name" and the fallback pattern.

    Patterns:
    - Pattern 1 (if product found): I came across [Company in Title Case] and your [food/beverage product in all lowercase], and I thought it'd be great to connect!
    - Pattern 2 (if no product): I came across [Company in Title Case], and I thought it'd be great to connect!

    Examples:
    - Input: Excel Juice Corp, products: cold-pressed juices → Output: I came across Excel Juice and your cold-pressed juices, and I thought it'd be great to connect!
    - Input: Steel Drinks Co., products: none → Output: I came across Steel Drinks, and I thought it'd be great to connect!
    - Input: Lily's Superfoods, products: acai bowls → Output: I came across Lily's and your acai bowls, and I thought it'd be great to connect!

    Now, write the icebreaker exactly in one sentence.
    """
    
    # Retry mechanism
    for attempt in range(0, 3): 
        response = model.generate_content(prompt, request_options={"timeout": 15})
        if response and response.text:
            return response.text.strip()
        else:
            return ""

# Generates the requested email
def generate_email(first_name, icebreaker, email_type, industry): # Be sure to comply with marketing email regulations like CAN-SPAM
    match industry:
        case "1":  # Food & Beverage
            match email_type:
                case "cold":
                    return f"""
                    <html>
                    <body>
                        <p>Hi {first_name},</p>
                        <p>{icebreaker}</p>
                        <p>I'm Josh, the founder and CEO of Food Inc. I've attached a description of our products below 
                        just so you have it on hand to check out. We offer free shipping and premium discounts on our website: 
                        <a href="https://www.yourwebsite.com">https://www.yourwebsite.com</a>.</p>
                        <p>
                            All the best,<br>
                            <p>Josh</p>
                            <br>
                            <span style="font-weight:bold;">Josh T</span> | <span style="font-weight:bold; color:blue">Food Inc.</span><br>
                            <span>CEO | </span> <a href="https://www.yourwebsite.com" style="font-weight:bold; color:blue;">yourwebsite.com</a><br>
                            <p>12345 Maple Street. Los Angeles, CA 90001</p><br>
                            <img src="cid:profileimage" width="98" style="margin-top:10px; border-radius:8px;"><br>
                        </p>
                        <p>P.S. I won’t flood your inbox, but reply with ‘unsubscribe’ if you’d prefer no more emails from me.</p>
                    </body>
                    </html>
                    """
                case "followup":
                    return f"""
                    <html>
                    <body>
                        <p>Hi {first_name},</p>
                        <p>Just wanted to follow up on my previous email, I understand things can get busy.</p>
                        <p>We’ve been producing quality products for the past 30 years. So, I thought it might be 
                        worthwhile to see if we would be a good fit for your business. </p>
                        <p>
                            All the best,<br>
                            <p>Josh</p>
                            <br>
                            <span style="font-weight:bold;">Josh T</span> | <span style="font-weight:bold; color:blue">Food Inc.</span><br>
                            <span>CEO | </span> <a href="https://www.yourwebsite.com" style="font-weight:bold; color:blue;">yourwebsite.com</a><br>
                            <p>12345 Maple Street. Los Angeles, CA 90001</p><br>
                            <img src="cid:profileimage" width="98" style="margin-top:10px; border-radius:8px;"><br>
                        </p>
                        <p>"P.S. I won’t flood your inbox, but reply with ‘unsubscribe’ if you’d prefer no more emails from me."</p>
                    </body>
                    </html>
                    """
        case "2":  # Sports Teams & Organizations
            match email_type:
                case "cold":  
                    return f"""
                    <html>
                    <body> 
                        <p>Hi {first_name},</p>
                        <p> I would love to send you free samples of our product for your athletes to try out!</p>
                        <p> Check out our attached recommendation from the lead dietition in Los Angeles.</p>
                        <p>
                            All the best,<br>
                            <p>Josh</p>
                            <br>
                            <span style="font-weight:bold;">Josh T</span> | <span style="font-weight:bold; color:blue">Food Inc.</span><br>
                            <span>CEO | </span> <a href="https://www.yourwebsite.com" style="font-weight:bold; color:blue;">yourwebsite.com</a><br>
                            <p>12345 Maple Street. Los Angeles, CA 90001</p><br>
                            <img src="cid:profileimage" width="98" style="margin-top:10px; border-radius:8px;"><br>
                        </p>
                        <p>"P.S. I won’t flood your inbox, but reply with ‘unsubscribe’ if you’d prefer no more emails from me."</p>
                    </body>
                    </html>
                    """ 
                case "followup":
                    return f"""
                    <html>
                    <body>
                        <p>Hi {first_name},</p>
                        <p>Just wanted to follow up on my previous email, I understand things can get busy.</p>
                        <p>We’ve been producing quality products for the past 30 years. So, I thought it might be 
                        worthwhile to see if our organics would be a good fit for your endeavors. </p>
                        <p>
                            All the best,<br>
                            <p>Josh</p>
                            <br>
                            <span style="font-weight:bold;">Josh T</span> | <span style="font-weight:bold; color:blue">Food Inc.</span><br>
                            <span>CEO | </span> <a href="https://www.yourwebsite.com" style="font-weight:bold; color:blue;">yourwebsite.com</a><br>
                            <p>12345 Maple Street. Los Angeles, CA 90001</p><br>
                            <img src="cid:profileimage" width="98" style="margin-top:10px; border-radius:8px;"><br>
                        </p>
                        <p>"P.S. I won’t flood your inbox, but reply with ‘unsubscribe’ if you’d prefer no more emails from me."</p>
                    </body>
                    </html>
                    """

# Main Program Loop
if __name__ == "__main__":
    # Update mail sent count
    with open("sent_data.txt", "r") as file:
        line = file.readline()
        sent_today_str, sent_total_str, date = line.strip().split(", ")
        sent_today = int(sent_today_str)
        sent_total = int(sent_total_str)
    if date != datetime.now().strftime("%m-%d-%Y"):
        with open("sent_data.txt", "w") as file:
            file.write("0, " + str(sent_total) + ", " + datetime.now().strftime("%m-%d-%Y"))
        sent_today = 0
        date = datetime.now().strftime("%m-%d-%Y")
    print() # SUPER COOL WELCOME GRAPHIC!!
    print(r'''
     ███████████████╗    ███╗   ███╗ █████╗ ██╗██╗     
     ╚═══██╔═══██╔══╝    ████╗ ████║██╔══██╗██║██║     
         ██║   ██║       ██╔████╔██║███████║██║██║     
         ██║   ██║       ██║╚██╔╝██║██╔══██║██║██║     
     ██████║   ██║       ██║ ╚═╝ ██║██║  ██║██║███████╗
     ╚═════╝   ╚═╝       ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝
            W e l c o m e    t o    J T m a i l        
    ''')
    # Main UI Loop
    while True:
        # if not check_license(): # Comment out for personal use
        #     print()
        #     print("Error: Incorrect email or subscription expired. Please contact support.")
        #     time.sleep(5)
        #     exit()
        print()
        print("Total Emails Sent: " + str(sent_total))
        print("Sent Today: " + str(sent_today))
        print()
        print("Type a number and press (Enter) to select an option:")
        print("[1] Send Cold Emails")
        print("[2] Send Follow-Up Emails")
        print("[3] Add to Do-Not-Contact List")
        print("[4] Exit JTmail")
        print()
        choice = input(">>> ").strip()

        email_type = ""
        match choice:
            case "1": # Cold Emails
                print()
                data = read_spreadsheet()
                if data is not None:
                    sent_today, sent_total = send_email(data, sent_today, sent_total, email_type = "cold")
                else:
                    print("No data to send emails.")
            case "2": # Follow-Up Emails
                print()
                data = pd.read_csv('sent.csv')
                if data is not None:
                    sent_today, sent_total = send_email(data, sent_today, sent_total, email_type = "followup")
                else:
                    print("No data to send emails.")
            case "3": # Add to Do-Not-Contact List
                entry = input("Enter the name or company to add to the Do-Not-Contact list: ").strip()
                if entry:
                    with open('do_not_contact.txt', 'a', encoding='utf-8') as f:
                        f.write(entry + '\n')
                    print(f"Added '{entry}' to the Do-Not-Contact list.")
                else:
                    print("No entry provided.")
            case "4": # Exit Program
                print()
                print("Exiting JTmail. Have a great day!")
                time.sleep(2)
                exit()
            case _:
                print("Invalid choice. Please try again.")
                continue         
