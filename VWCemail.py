#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Pip Installs
#!pip install pywin32


# In[ ]:


import pandas as pd
from IPython.display import display
import win32com.client as win32


# In[ ]:


# Create a English-French dictionary
engDict = {
    "Field 1 - Office of the Chief Statistician": "Secteur 1 - Bureau du Statisticien en Chef",
    "Field 3 - Corporate Strategy and Management": "Secteur 3 - Stratégies et Gestion Intégrées",
    "Field 4 - Strategic Engagement": "Secteur 4 - Engagement Stratégique",
    "Field 5 - Economics Statistics": "Secteur 5 - Statistiques Économique",
    "Field 6 - Strategic Data Management, Methods, and Analysis": "Secteur 6 - Gestion Stratégique des Données, Méthodes et Analyse",
    "Field 7 - Census, Regional Services, and Operations": "Secteur 7 - Recensement, Services Régionaux, et Opérations",
    "Field 8 - Social Health and Labour Statistics": "Secteur 8 - Statistiques Sociale, de la Santé et du Travail",
    "Field 9 - Digital Solutions": "Secteur 9 - Solutions Numériques"
}


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

# Display
display(df)


# In[ ]:


"""Automated Emails for Matched Groups"""
# For each pair in list of matches
for i in range(len(df)):

    # Body text of email (English)
    text = """Hello {},


You have been matched together for a Virtual Watercooler conversation. We recommend using MS 
Teams scheduled during regular business hours for a conversation of at about 10 minutes but it is up to 
you to decide how to proceed.

The group prefers to chat in {} in the {}. You work in {}, respectively.

{}

Please reach out to Innovation Coordinator / Coord de l'innovation 
(STATCAN) statcan.innovationcoordinator-coorddelinnovation.statcan@canada.ca with all of your 
feedback, questions and suggestions. Thank you for using the StatCan Virtual Watercooler.



Sincerely,

The StatCan Virtual Watercooler Team

Innovation Secretariat""".format(df.iloc[i]['name'].replace(",", " and"), # Names
                                 df.iloc[i]['language'], # Language preference
                                 df.iloc[i]['time'], # Time preference
                                 df.iloc[i]['field'].replace(",", " and"),
                                 df.iloc[i]['interests'] # Interests
                                ) # Field
    # columns = [matched, email, name, language, time, field, interests, langue, temps, interestsFR]

    # French translation of the email
    textFr = """Bonjour {},


Vous avez été jumelés pour une causerie virtuelle. Nous vous recommandons d’utiliser MS Teams 
pendant les heures normales de travail pour discuter environ 10 minutes, mais c’est à vous de décider 
de la manière de procéder.

Le groupe préfère discuter en {} dans {}. Vous travaillez dans {}, respectivement.

{}

Nous vous invitons à communiquer avec le coordonnateur de l’innovation de Statistique Canada 
(statcan.innovationcoordinator-coorddelinnovation.statcan@canada.ca) 
si vous avez des commentaires, des questions et des suggestions. Nous vous remercions de participer aux 
causeries virtuelles de Statistique Canada.


Bien cordialement,

L’Équipe des causeries virtuelles de Statistique Canada 

Secrétariat de l’innovation""".format(df.iloc[i]['name'].replace(",", " et"), # Names
                                      df.iloc[i]['langue'], # Language preference
                                      df.iloc[i]['temps'], # Time preference
                                      " et ".join([engDict.get(word,word) for word in (df.iloc[i]['field'].replace(",", "").split())]), # Field
                                      df.iloc[i]['interestsFR'] # Interests
                                     )

    # Final email message
    message = text + "\n\n\n" + textFr
    print(message)

    # The emails from each person in the pair
    recipients = df.iloc[i]['email'].replace(",", ";")

    # Send the emails
    #email(recipients, message)


# In[ ]:




