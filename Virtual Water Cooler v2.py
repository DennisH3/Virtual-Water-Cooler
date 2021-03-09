#!/usr/bin/env python
# coding: utf-8

# In[43]:


# Pip Installs
#!pip install google_trans_new

# Install spacy and the the large english model
#!pip install spacy
#!python -m spacy download en_core_web_lg


# In[44]:


"""Initialization"""

# Import statements
import pandas as pd
import os
from IPython.display import display

from google_trans_new import google_translator
translator = google_translator()

import spacy
import en_core_web_lg
nlp = en_core_web_lg.load()


# In[45]:


"""Dictionaries"""

# Create a French-English dictionary
frenchDict = {
    "Anglais": "English",
    "Français": "French",
    "Matin": "Morning",
    "Après-midi": "Afternoon",
    "Pas de préférence": "No preference",
    "Secteur 1 - Bureau du Statisticien en Chef": "Field 1 - Office of the Chief Statistician",
    "Secteur 3 - Stratégies et Gestion Intégrées": "Field 3 - Corporate Strategy and Management",
    "Secteur 4 - Engagement Stratégique": "Field 4 - Strategic Engagement",
    "Secteur 5 - Statistiques Économique": "Field 5 - Economics Statistics",
    "Secteur 6 - Gestion Stratégique des Données, Méthodes et Analyse": "Field 6 - Strategic Data Management, Methods, and Analysis",
    "Secteur 7 - Recensement, Services Régionaux, et Opérations": "Field 7 - Census, Regional Services, and Operations",
    "Secteur 8 - Statistiques Sociale, de la Santé et du Travail": "Field 8 - Social Health and Labour Statistics",
    "Secteur 9 - Solutions Numériques": "Field 9 - Digital Solutions",
    "Oui": "Yes",
    "Non": "No"
}

# Column names
colname_dict = {
    "Please enter your @canada.ca email.": "email",
    "What is your preferred name?": "name",
    "What language would you like to converse in?": "language",
    "When would you like to chat?": "time",
    "Which field are you in?": "field",
    "Do you want to be matched ONLY WITHIN your field?": "field_match",
    "What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.": "interests",
}


# In[46]:


"""Helper Functions"""


"""
Desc:
    Function that checks for pair matching
    Params: a, b (row in pandas DataFrame)
    Output: boolean
"""
def match(a, b):
    if a['language'] != b['language'] and "No preference" not in [a['language'], b['language']]:
        return False
    elif a['time'] != b['time'] and "No preference" not in [a['time'], b['time']]:
        return False
    elif a['field'] != b['field'] and "Yes" in [a['field_match'], b['field_match']]:
        return False
    else:
        a['matched'] = "Yes"
        b['matched'] = "Yes"
        return True


"""
Desc:
    Function that picks the language.
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def language(pair):
    # If the first person or the second person said English
    if ((pair.iat[0, 2] == "English") | (pair.iat[1, 2] == "English")):

        # Return English
        return "English"

    # Else if the first person or second person said French
    elif ((pair.iat[0, 2] == "French") | (pair.iat[1, 2] == "French")):

        # Return French
        return "French"

    # Else it was No preference
    else:
        return "either English or French"


"""
Desc:
    Function that picks the time.
    Params: pair (pandas DataFrame)
    Output: the time (string)
"""
def time(pair):
    # If the first person or the second person said Morning
    if ((pair.iat[0, 3] == "Morning") | (pair.iat[1, 3] == "Morning")):

        # Return Morning
        return "Morning"

    # Else if the first person or second person said Afternoon
    elif ((pair.iat[0, 3] == "Afternoon") | (pair.iat[1, 3] == "Afternoon")):

        # Return Afternoon
        return "Afternoon"

    # Else it was No preference
    else:
        return "Morning or Afternoon"


"""
Desc:
    Function that picks the language (French).
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def langue(pair):
     # If the first person or the second person said English
    if ((pair.iat[0, 2] == "English") | (pair.iat[1, 2] == "English")):

        # Return Anglais
        return "Anglais"

    # Else if the first person or second person said French
    elif ((pair.iat[0, 2] == "French") | (pair.iat[1, 2] == "French")):

        # Return Français
        return "Français"

    # Else it was No preference
    else:
        return "soit Anglais ou Français"


"""
Desc:
    Function that picks the time (French).
    Params: pair (pandas DataFrame)
    Output: the language (string)
"""
def temps(pair):
    # If the first person or the second person said Morning
    if ((pair.iat[0, 3] == "Morning") | (pair.iat[1, 3] == "Morning")):

        # Return la Matinée
        return "la Matinée"

    # Else if the first person or second person said Afternoon
    elif ((pair.iat[0, 3] == "Afternoon") | (pair.iat[1, 3] == "Afternoon")):

       # Return l'Après-midi
        return "l'Après-midi"

    # Else it was No preference
    else:
        return "la Matinée ou l'Après-midi"


# In[47]:


"""Load the Data"""

# Read the English response csv file skipping the column header row.
dfEng = pd.read_csv("EngResponsesBeta.csv")

# Read the French response csv skipping the column header row
dfFr = pd.read_csv("FrResponsesBeta.csv")


# In[48]:


"""Translate from French to English"""

# Make them have the same header as the English ones
dfFr.columns = dfEng.columns

# Map the values from columns 3-6 to their dictionary values (Translate French to English)
dfFr.iloc[:, 3:7] = dfFr.iloc[:, 3:7].applymap(frenchDict.get)

# Combine the English dataframe with the French dataframe (on horizontal axis) into one dataframe
df = dfEng.append(dfFr, ignore_index=True)

# Remove the first column (Time Stamp column)
df = df.drop(df.columns[[0]], axis=1)

# Convert the Interest column into String type
df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'] = (df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'].astype(str))

# Translate each list of interests into English
for i in range(len(df)):

    # If the interest is in French
    if (translator.detect(df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i])[0] == 'fr'):

        # Translate it
        df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i] = translator.translate(df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i])

# Create the matched column and fill the column with No
df["matched"] = "No"

# Rename the columns and drop duplicates
df = df.rename(colname_dict, axis=1).drop_duplicates(subset=['email'], keep='last')

# Display
display(df)


# In[49]:


"""Match Making Algorithm"""

"""
Desc:
    Function that makes pairs.
    Params: df (pandas Data Frame)
    Output: match_successes (list of pandas Data Frame), unmatched_df (pandas Data Frame)
"""
def matchMaking(df):

    # Transform the dataframe into a list of dictionaries
    records = df.to_dict('records')
    total_records = len(records)

    # Prepare lists to store matches and non-matches
    match_successes = []
    match_failures = []

    # Main loop
    while len(records) > 1:
        for i, n in enumerate(records):
            if i > 0:
                if match(records[0], n) is True:
                    match_successes.append([records.pop(i), records.pop(0)])
                    print(f"Match! records left: {len(records)}/{total_records}")
                    break
            if i+1 == len(records):
                match_failures.append(records.pop(0))
                print("Cannot match a record...")

    # Add the last record to the match failures
    if len(records) > 0:
        match_failures.append(records.pop(0))

    print(f"Matching complete!")
    print(f"Number of matched records: {2*len(match_successes)}")
    print(f"Number of unmatched records: {len(match_failures)}")

    # Create a dataframe for non-matches
    unmatched_df = pd.DataFrame(match_failures)

    return match_successes, unmatched_df


# In[91]:


# Function to find common interests
def common_interests(df):

    # If any of the pairs containt NaN (since NaN is a float)
    if isinstance(df['interests'][0], float) and isinstance(df['interests'][1], float):
        return "Both of you did not include your interests"
    elif isinstance(df['interests'][0], float):
        return "Your common interests are: " + str(df['interests'][1])
    elif isinstance(df['interests'][1], float):
        return "Your common interests are: " + str(df['interests'][0])
    else:
        # Split the string of interests by comma to create a list of interests
        i1 = df['interests'][0].split(",")
        i2 = df['interests'][1].split(",")

        # Lemmatize the words in both lists
        # Treat each word as a "document"
        # If the word is a pronoun, leave it
        for i in range(len(i1)):
            if nlp(i1[i])[0].lemma_ == "-PRON-":
                i1[i] = nlp(i1[i]).text
            else:
                i1[i] = nlp(i1[i])[0].lemma_

        for j in range(len(i2)):
            if nlp(i2[j])[0].lemma_ == "-PRON-":
                i2[j] = nlp(i2[j]).text
            else:
                i2[j] = nlp(i2[j])[0].lemma_

        # Empty list to store common interests
        commonInterests = []

        # Find the simmilarity of all the words
        # from the shorter list compared to the longer list
        if (len(i1) <= len(i2)):
            for i in range(len(i1)):
                for j in range(len(i2)):

                    # Skip blank characters
                    if (nlp(i1[i]).vector_norm and nlp(i2[j]).vector_norm):
                        if (nlp(i1[i]).similarity(nlp(i2[j])) >= 0.6):
                            commonInterests.append(i1[i])
        else:
            for i in range(len(i2)):
                for j in range(len(i1)):

                    # Skip blank characters
                    if (nlp(i1[j]).vector_norm and nlp(i2[i]).vector_norm):
                        if (nlp(i2[i]).similarity(nlp(i1[j])) >= 0.6):
                            commonInterests.append(i2[i])

        # Remove duplicate words
        commonInterests = set(commonInterests)

        # If the list of common interests is not empty
        if (len(commonInterests) > 0):
            return "Your common interests are: " + ','.join(commonInterests)
        else:
            # Else the commonInterests list is empty, combine their interests
            commonInterests = (i1 + i2)

            # Remove blank spaces
            commonInterests[:] = [x for x in commonInterests if x.strip()]

            return "Your common interests are: " + ','.join(commonInterests)


# In[92]:


"""Match Making"""

# Note: The first time the code is run prevVWCData.csv should just be an empty CSV file

# Check if prevVWCData.csv is empty or not
# If it is not empty
if (os.stat("prevVWCData.csv").st_size != 0):

    # Read the previous file and rename the columns
    pData = pd.read_csv("prevVWCData.csv").rename(colname_dict, axis=1)

    # Check if there were new entries
    if (len(pData) == len(df)):

        # There were no new entries
        print("There were no new entries")

    # There are new entries
    else:

        # Replace the rows in df with matched = Yes
        cols = list(df.columns)
        df.loc[df.email.isin(pData.email), cols] = pData[cols]

        # Drop the rows where matched = Yes
        df = df.drop(df[df.matched == "Yes"].index)

        # Store the list of matches and no matches
        matches = matchMaking(df)[0]
        noMatches = matchMaking(df)[1]

        # Empty list to store common interests
        ci = []

        # Emmpty list to store common interests in French
        ciFR = []

        # Convert each pair into data frame
        for i in range(len(matches)):
            matches[i] = pd.DataFrame(matches[i])

            # Fill ci
            ci.append(common_interests(matches[i]))

        # Translate each list of interests into French
        for i in ci:
            ciFR.append(translator.translate(i, lang_tgt="fr"))

        # Find common language preference
        lang = [language(pair) for pair in matches]

        # Find common time preference
        t = [time(pair) for pair in matches]

        # Translate common language preference
        langFR = [langue(pair) for pair in matches]

        # Translate common time preference
        tFR = [temps(pair) for pair in matches]

        # List to store merged pairs
        mergeDFs = []

        # Join the pairs into one row
        for pair in matches:
            mergeDFs.append(pair.groupby('matched')[['email','name','language', 'time', 'field', 'interests']].agg(', '.join))

        # Combine them into one data frame
        emailDF = pd.concat(mergeDFs).reset_index()

        # Update language, time, interests columns, and add French versions
        emailDF["language"] = lang
        emailDF["time"] = t
        emailDF["interests"] = ci
        emailDF["langue"] = langFR
        emailDF["temps"] = tFR
        emailDF["interestsFR"] = ciFR

        """Export the result to be used for Email program"""
        emailDF.to_csv("./emailVWC.csv", index=False)
        # Exports a csv file with columns = [matched, email, name, language,
        # time, field, interests, langue, temps, interestsFR]

        """Update prevVWCData.csv"""
        # Add the noMatches data frame into the list of data frames
        matches.append(noMatches)

        print(matches)

        # Concatenate all the data frames into one
        newDF = pd.concat(matches)

        # Update pData
        cols = list(pData.columns)
        pData.loc[pData.email.isin(newDF.email), cols] = newDF[cols]

        # Reverse the column naming
        pData = pData.rename({v:k for k, v in colname_dict.items()}, axis=1)

        # Overwrite prevVWCData.csv
        pData.to_csv("./prevVWCData.csv", index=False)

# Else, prevVWC.csv is empty, perform the matchmaking
else:
    # Store the list of matches and no matches
    matches = matchMaking(df)[0]
    noMatches = matchMaking(df)[1]

    # Empty list to store common interests
    ci = []

    # Emmpty list to store common interests in French
    ciFR = []

    # Convert each pair into data frame
    for i in range(len(matches)):
        matches[i] = pd.DataFrame(matches[i])

        # Fill ci
        ci.append(common_interests(matches[i]))

    # Translate each list of interests into French
    for i in ci:
        ciFR.append(translator.translate(i, lang_tgt="fr"))

    # Find common language preference
    lang = [language(pair) for pair in matches]

    # Find common time preference
    t = [time(pair) for pair in matches]

    # Translate common language preference
    langFR = [langue(pair) for pair in matches]

    # Translate common time preference
    tFR = [temps(pair) for pair in matches]

    # List to store merged pairs
    mergeDFs = []

    # Join the pairs into one row
    for pair in matches:
        mergeDFs.append(pair.groupby('matched')[['email','name','language', 'time', 'field', 'interests']].agg(', '.join).reset_index())

    # Combine them into one data frame
    emailDF = pd.concat(mergeDFs)

    # Update language, time, interests columns, and add French versions
    emailDF["language"] = lang
    emailDF["time"] = t
    emailDF["interests"] = ci
    emailDF["langue"] = langFR
    emailDF["temps"] = tFR
    emailDF["interestsFR"] = ciFR

    """Export the result to be used for Email program"""
    emailDF.to_csv("./emailVWC.csv", index=False)
    # Exports a csv file with columns = [matched, email, name, language,
    # time, field, interests, langue, temps, interestsFR]

    """Update prevVWCData.csv"""
    # Add the noMatches data frame into the list of data frames
    matches.append(noMatches)

    # Concatenate all the data frames into one
    newDF = pd.concat(matches)

    # Reverse the column naming
    newDF = newDF.rename({v:k for k, v in colname_dict.items()}, axis=1)

    # Display
    # display(newDF)

    # Overwrite prevVWCData.csv
    newDF.to_csv("./prevVWCData.csv", index=False)


# In[56]:


# Convert jupyter notebook to python script
get_ipython().system('jupyter nbconvert --to script "Virtual Water Cooler v2.ipynb"')


# In[ ]:




