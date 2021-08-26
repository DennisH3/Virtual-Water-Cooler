#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Pip Installs
get_ipython().system('pip install google_trans_new')

# Install spacy and the the large english model
get_ipython().system('pip install spacy')
get_ipython().system('python -m spacy download en_core_web_lg')


# In[ ]:


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


# In[ ]:


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
    "Hebdomadaire": "Weekly",
    "Une seule fois, je me réinscrirai si je veux un nouveau match.": "Only once, I will sign up again if I want a new match."
}

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

# Column names
colname_dict = {
    "Please enter your @canada.ca email.": "email",
    "What is your preferred name?": "name",
    "What language would you like to converse in?": "language",
    "When would you like to chat?": "time",
    "Which field are you in?": "field",
    "How often would you liked to be matched with a new colleague?": "freq",
    "What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.": "interests",
}


# In[ ]:


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
    else:
        a['matched'] = 1
        b['matched'] = 1
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


# In[ ]:


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


# In[ ]:


"""
Desc:
    Function that finds common interests.
    Params: df (pandas Data Frame)
    Output: a string that enumerates a pair of user's common interests
"""
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


# In[ ]:


"""Load the Data"""

# Read the English response csv file skipping the column header row.
dfEng = pd.read_csv("EngResponsesBeta.csv")

# Read the French response csv skipping the column header row
dfFr = pd.read_csv("FrResponsesBeta.csv")


# In[ ]:


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
for i in range(len(df.index)):

    # If the interest is in French
    if (translator.detect(df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i])[0] == 'fr'):

        # Translate it
        df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i] = translator.translate(df['What are your interests? Please list your interests separated by a comma with no spaces. I.e. cooking,drawing,etc.'][i])

# Create the matched column and fill the column with 0
df["matched"] = 0

# Rename the columns and drop duplicates, keeping only most recent submission
df = df.rename(colname_dict, axis=1).drop_duplicates(subset=['email'], keep='last')

# Create adjacency matrix (for random sampling without replacement to keep track of prevMatches so matched people don't form the same pair)
# adj_matrix = pd.DataFrame(np.zeros(shape=(len(df.index),len(df.index))), columns=df['name'].unique())

# Check if people who only wanted to be matched one time were matched
if (os.stat("emailVWC.csv").st_size != 0):

    # Read the prevData file
    prevDF = pd.read_csv("emailVWC.csv")

    # Select only the common columns
    prevDF = prevDF[["email", "name", "language", "time", "field", "freq", "interests", "matched"]]

    # Merge with the prevDF and keep values from prevDF
    df = df.merge(prevDF, how='right')

    # Exclude people who have been matched and only wanted to be matched once
    df = df.drop(df[(df["matched"] == 1) & (df["freq"] == "Only once, I will sign up again if I want a new match.")].index)

    # Reset all matches to 0
    df["matched"] = 0

# Shuffle and reset index
df = df.sample(frac=1).reset_index(drop=True)

# Display
display(df)


# In[ ]:


"""Match Making"""

# Note: The first time the code is run emailVWC.csv should be an empty file

 # Store the list of matches and no matches
matches = matchMaking(df)[0]
noMatches = matchMaking(df)[1]

# Convert each pair into data frame
for i in range(len(matches)):

    # Convert each pair to a data frame
    matches[i] = pd.DataFrame(matches[i])

    # Make a column for common interests
    matches[i]["comInterests"] = common_interests(matches[i])

    # Make a column for common interests in French
    matches[i]["comInterestsFR"] = translator.translate(matches[i]["comInterests"][0], lang_tgt="fr")

    # Make a column for language preference
    matches[i]["lang"] = language(matches[i])

    # Make a column for language preference
    matches[i]["langFR"] = langue(matches[i])

    # Make a column for time preference
    matches[i]["t"] = time(matches[i])

    # Make a column for time preference
    matches[i]["tFR"] = temps(matches[i])

    # Print to check
    display(matches[i])

# Combine matches and noMatches into one data frame
emailDF = pd.concat([pd.concat(matches, ignore_index=True), noMatches], ignore_index=True)

# Translate field into french
emailDF["fieldFR"] = emailDF["field"].map(engDict.get)

# Convert all columns to string
emailDF = emailDF.astype(str)

# Convert the matched column to integer
emailDF["matched"] = emailDF["matched"].astype(int)

# Display
display(emailDF)

"""Export the result to be used for Email program"""
# Exports a csv file with columns = [email,name,language,time,field,freq,interest, cominterests,matched,cominterestsFR,lang,langFR,t,tFR,fieldFR]
emailDF.to_csv("./emailVWC.csv", index=False)


# In[ ]:


# Convert jupyter notebook to python script
#!jupyter nbconvert --to script "Virtual Water Cooler v2.ipynb"

