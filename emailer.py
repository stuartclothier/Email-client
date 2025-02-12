import random
from datetime import datetime as dt
import pandas as pd
import win32com.client as win32




greeting_list = ["Hello!", "Hi there!", "Hey,\n\nHow are you?", "Greetings!", "Greetings!",
                 "Howdy!", "Nice to see you!", "Dear Sir and/or Madam",
                 "Hello,\n\nI hope you're well",
                 "Hi,\n\nHow have you been?",
                 "Hey,\n\nWhat's up?", "Hello,\n\nHow's your day going?",
                 "Hi,\n\nIt's great to see you again!",
                 "Hey,\n\nHow's everything with you?",
                 "Hello!", "Hi there!",
                 "Hello!", "Hi there!",
                 "Hello!", "Hi there!",
                 "Hi,\n\nHope you're doing well.",
                 "Hi,\n\nI hope you're having a good day.",
                 "Hi,\n\nI hope you're having a good day.",
                 "Hi,\n\nI hope you're having a good day.",
                 "Greetings!",
                 "Hey,\n\nIt's good to be in touch with you!",  
                 "Most esteemed recipient,\n\nAccept my heartfelt greetings and " +
                 "well wishes. I hope this message finds thee in good health and high spirits."]

item_list = ["The System has successfully recorded this item.",
             "We have added this item to the System.",
             "This item has been registered in the System.",
             "You can now find this item within the System.",
             "The System now contains information about this item.",
             "Successfully logged: This item is now part of the System.",
             "Confirmation: This item has been integrated into the System.",
              "Status update: The System now reflects the presence of this item.",
              "Recorded: This item is officially in the System.",
             "Updated: The System now holds data about this item.",
             "We have added this item to the System.",
             "We have added this item to the System.",
             "This item has been added to the System.", 
             "This item has been added to the System.",
             "This item has been added to the System.",
             "This item has been added to the System.", 
             "This item has been added to the System.",
             "This item has been added to the System.",
             "New addition gleams,\nSystem welcomes with delight,\nIn ones and zeroes."]

items_list = ["The System has successfully recorded these items:",
              "We have added these items to the System:",
              "These items have been registered in the System:",
              "You can now find these items within the System:",
              "Successfully logged: These items are now part of the System:",
              "These items have been added to the System:",
              "These items have been added to the System:",
              "These items have been added to the System:",
              "These items have been added to the System:",
              "These items have been added to the System:",
              "These items have been added to the System:",
              "We have added these items to the System:",
              "We have added these items to the System:",
              "We have added these items to the System:",
              "In accordance with the esteemed practice," +
               " these esteemed possessions are now catalogued" +
               " in the grand Chronicles of our venerable Institution:"]


goodbye_list = ["Kind Regards", "Thanks", "Best regards", "Sincerely", 
                "Warm regards", "Best wishes", "Regards", "Yours sincerely",
                  "Thank you", "With gratitude", "Much appreciated",
                "Many thanks","Take care", "Have a great day", "All the best",
                "Take care and talk soon", "Have a wonderful week", "With best regards",
                "Kind Regards", "Thanks", "Best regards", "Kind Regards", "Thanks",
                "Best regards", "Have a great day",
                "Accept my parting regards, esteemed colleague. May thy endeavours " +
                "be fruitful, and may we meet again in harmonious fellowship."]


# uncomment/comment below lines to toggle the name of user
# name_list = ["Ethelbert Scanner III", "Ethelbert", "Berty"]
name_list = ["Stuart", "Stuart", "Stu"]


# @Gooey
def draft_and_open_email():
    # parser = ArgumentParser(...)
    # Create an instance of the Outlook application

    excel_file_path = 'test_email.xlsx'
    sheet_name = 'Sheet1'  # Modify this if your sheet name is different

    # Read the Excel data using pandas
    data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    # Group rows by the 'To_Email' column and concatenate subjects and bodies
    data['Subject'] = data['Subject'].astype(str)
    data["concat"] = data['Subject']+" - "+data["Desc"]
    grouped_data = data.groupby('To_Email')['concat'].apply(
        ' | '.join).reset_index()

    for _, row in grouped_data.iterrows():
        outlook = win32.Dispatch("Outlook.Application")
        # 0 represents olMailItem (email item)
        mail_item = outlook.CreateItem(0)
        mail_item.SentOnBehalfOfName = "products@opchealth.com.au"
        mail_item.To = row['To_Email']

        # uncomment/comment below line to toggle cc to products email
        # mail_item.CC = "products@opchealth.com.au"

        greeting = random.choice(greeting_list)
        goodbye = random.choice(goodbye_list)
        if ' | ' in row['concat']:
            item = random.choice(items_list)+"\n"
            for _, row2 in data.loc[data['To_Email'] == row['To_Email']].iterrows():
                item = item + "\n"+row2['Subject']+" - "+row2["Desc"]
            mail_item.Subject = "Item additions " + dt.now().strftime("%Y-%m-%d")

        else:
            item = random.choice(item_list)
            mail_item.Subject = row['concat']

        # uncomment/comment below line to toggle random name
        # name = random.choice(name_list)

        # +"\n\n"+name commented out to supress name in email body
        mail_item.Body = greeting + "\n\n"+item+"\n\n"+goodbye  # +"\n\n"+name

        mail_item.Display()


draft_and_open_email()
