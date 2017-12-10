#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: aymeric
"""

import string   
import re
import pandas as pd
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from gensim.models import Word2Vec
import json
from nltk.stem.wordnet import WordNetLemmatizer
#from autocorrect import spell


xl = pd.ExcelFile("Word-cloud-contributions-Examples.xls")  #load the excel file
names=xl.sheet_names[1:]       #get the names for each sheet except the introduction one

#we use both French and English stop words because they are not included in the words of the other language
stop_words_english=set(stopwords.words('english'))  #lists of English words that have to be removed from data
stop_words_french=set(stopwords.words('french'))    #lists of French words that have to be removed from data

punctuation=string.punctuation   #list of punctuation sign that must be removed from data
punctuation+="''"+"’’"+"``"+"n't"+"'m" +"'s"+"'d"+"'re"        #we add the abreviation that must be removed
dictionnary={}   #to store the words and numbers of occurence for each word
data=[]     #data for the word2vec


for sheet_name in names:
    table=xl.parse(sheet_name)          #get the table
    for k in range(len(table)):
        row=table["Comment"][k]         #get the row
        row_data=row.strip().split(".")  #if there is a point  between two words that 
                                        #means thre is 2 sentences                         
        for i in range(len(row_data)):
            line=row_data[i].lower()            #all the words are rewritten in lower letters to avoid duplicates
            token=nltk.word_tokenize(line)      #remove useless words
            cleaned_data=[]
            
            for word in token:
                word=word.strip()       #remove the extra spaces
                #if the word is not in stop words or punctuatation
                if not((word in stop_words_english)or(word in stop_words_french)or(word in punctuation)):
                    cleaned_data.append(word)
                    if len(word.strip())<=1:    #if the word has less than one letter it is removed
                        continue
                    else:
                        try:
                            dictionnary[word]+=1
                        except:
                            dictionnary[word]=1
            if len(cleaned_data)>0 :   
                data.append(cleaned_data)
 
    
#for verbs we change it to the present version of the verb (ex:gives=>give   working=>work)
deleted_keys=[]
keys_to_add=[]
wnl = WordNetLemmatizer()

for key in dictionnary:
    correction=wnl.lemmatize(key,"v")
    if correction!=key:
        if key not in deleted_keys:
            deleted_keys.append(key)
        try:
            dictionnary[correction]+=dictionnary[key]
        except:
            keys_to_add.append([dictionnary[key],correction])
  
          
#we add and create new keys after the itertion in the dictionary not to change the size during the loop
#create new keys
for k in range(len(keys_to_add)):
    key=keys_to_add[k][1]
    value=keys_to_add[k][0]
    dictionnary[key]=value

#remove the wrong keys
for key in deleted_keys:
    del dictionnary[key]
print(len(deleted_keys)-len(keys_to_add), " keys were deleted by changing verbs")     


 #make sure we remove the plurals  
deleted_keys2=[] 
keys_to_add2=[]
for key in dictionnary:
    correction=wnl.lemmatize(key)
    if correction!=key:
        if key not in deleted_keys:
            deleted_keys2.append(key)
        try:
            dictionnary[correction]+=dictionnary[key]
        except:
            keys_to_add2.append([dictionnary[key],correction]) 
            
for k in range(len(keys_to_add2)):
    key=keys_to_add2[k][1]
    value=keys_to_add2[k][0]
    dictionnary[key]=value            
        
for key in deleted_keys2:
    del dictionnary[key]      
print(len(deleted_keys2)-len(keys_to_add2), "keys were deleted by removing plurals")
 
              
sorted_keys=sorted(dictionnary, key=lambda keys: dictionnary[keys])  #keys of the dictionnary sorted by increasing values 
print("The dictionnary has ",len(sorted_keys)," keys")



#save the data in files
#with open('data.json', 'w') as outfile:
#    json.dump(dictionnary, outfile)

#with open('keys.txt', 'w') as outfile:
#    json.dump(sorted_keys, outfile)    


        

#the word2vec model was made when i thought I could try to put close to each other
#words with the same meaning but as the idea of the word cloud was to put the most used words first
#I couldn't do it

# train model
model = Word2Vec(data, min_count=1)
# display the most similar words from a word
#print(model.most_similar("skype",topn=5))

##display the vector
#print(model['compliance'])

