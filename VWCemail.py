
# coding: utf-8

# In[ ]:


# Pip Installs
#!pip install pywin32


# In[ ]:


import pandas as pd
from IPython.display import display
import win32com.client as win32
import time


# In[ ]:


"""
Desc:
    Function that sends an email.
    Params: recipients (string), subject (string), text (string)
    Output: An email
    Note: you must be logged onto your Outlook 2013 account first before this will run
"""
def email(recipients, text, profilename="Outlook 2013"):
    oa = win32.Dispatch("Outlook.Application")

    Msg = oa.CreateItem(0)
    Msg.To = recipients

    Msg.Subject = "Virtual Water Cooler"
    Msg.Body = text

    Msg.Display()
    # Msg.Send()


# In[ ]:


"""Load the data"""

df = pd.read_csv("emailVWC.csv")

# Filter the data for people who have been matched
df = df[df["matched"] == 1]

# Display
display(df)


# In[ ]:


"""Automated Emails for Matched Groups"""

# For each pair in list of matches
for i in range(0, len(df.index)-1, 2):

    # Body text of email (English)
    text = """Hello {} and {},


You have been matched together for a Virtual Watercooler conversation. We recommend using MS 
Teams scheduled during regular business hours for a conversation of at about 10 minutes, but it is up to 
you to decide how to proceed.

The group prefers to chat in {} in the {}. You work in {} and {}, respectively.

{}

Please reach out to [name] <email> with all of your 
feedback, questions and suggestions. Thank you for using the StatCan Virtual Watercooler.



Sincerely,

The StatCan Virtual Watercooler Team

Innovation Secretariat""".format(df.iloc[i]['name'], # Name of Person 1
                                 df.iloc[i+1]['name'], # Name of Person 2
                                 df.iloc[i]['lang'], # Language preference
                                 df.iloc[i]['t'], # Time preference
                                 df.iloc[i]['field'], # Field of Person 1
                                 df.iloc[i+1]['field'], # Field of Person 2
                                 df.iloc[i]['comInterests'] # Common interests
                                )

    # French translation of the email
    textFr = """Bonjour {} et {},


Vous avez été jumelés pour une causerie virtuelle. Nous vous recommandons d’utiliser MS Teams 
pendant les heures normales de travail pour discuter environ 10 minutes, mais c’est à vous de décider 
de la manière de procéder.

Le groupe préfère discuter en {} dans {}. Vous travaillez dans {} et {}, respectivement.

{}

Nous vous invitons à communiquer avec [nom] <email>
si vous avez des commentaires, des questions et des suggestions. Nous vous remercions de participer aux 
causeries virtuelles de Statistique Canada.


Bien cordialement,

L’Équipe des causeries virtuelles de Statistique Canada 

Secrétariat de l’innovation""".format(df.iloc[i]['name'], # Name of Person 1
                                      df.iloc[i+1]['name'], # Name of Person 2
                                      df.iloc[i]['langFR'], # Language preference
                                      df.iloc[i]['tFR'], # Time preference
                                      df.iloc[i]['fieldFR'], # Field of Person 1
                                      df.iloc[i+1]['fieldFR'], # Field of Person 2
                                      df.iloc[i]['comInterestsFR'] # Common interests
                                     )

    # Final email message
    message = text + "\n\n\n" + textFr
    print(message)

    # The emails from each person in the pair
    recipients = df.iloc[i]['email'] + "; " + df.iloc[i+1]['email']
    print(recipients)
    
    # Send the emails
    email(recipients, message)
    
    # Wait 3 seconds before next email
    time.sleep(3)

