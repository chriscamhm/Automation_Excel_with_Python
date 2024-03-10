#!/usr/bin/env python
# coding: utf-8

# In[47]:


import pandas as pd
import warnings
warnings.filterwarnings('ignore')


# In[48]:


def normalize(s):
    replacements=(
    ("á","a"),
    ("é","e"),
    ("í","i"),
    ("ó","o"),
    ("ú","u"),
    )
    for a,b in replacements:
        s = s.replace(a,b).replace(a.upper(),b.upper())
    return s


# In[49]:


df_c=pd.read_excel('comentarios totales ecopetrol abiertos.xlsx', sheet_name='comenzar')
df_c


# In[50]:


df_d=pd.read_excel('comentarios totales ecopetrol abiertos.xlsx', sheet_name='dejar')
df_d


# In[51]:


df_c['dup'] = df_c.duplicated(subset=None, keep='first')
df_c = df_c[df_c['dup'] == False]
df_c


# In[52]:


terms1 = ['vida personal', 'familia', 'familiares', 'horario', 'tiempo libre', 'vida laboral', 'balance','descanso','reuniones','reunión','equilibrio']
df_c= df_c[df_c['Comentario'].str.contains('|'.join(terms1), na=False)]
df_c


# In[53]:


import nltk
import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize


example_sent = "Hola, soy Cristian y esto es la cosa"
stop_words = set(stopwords.words('spanish'))
word_tokens = word_tokenize(example_sent)

filtered_sentence = [w for w in word_tokens if not w in stop_words]
print(filtered_sentence)


# In[54]:


# Import the package
import stanza
stanza.download('es')
nlp = stanza.Pipeline('es')


# In[55]:


terms2 = ['vida personal', 'familia', 'familiares', 'horario', 'tiempo libre', 'vida laboral', 'balance','descanso','reuniones','reunión','equilibrio']
def preprocess(raw_text):
    
    stop_words = set(stopwords.words('spanish'))
    raw_text= normalize(raw_text)
    
    #regular expression keeping only letters 
    letters_only_text = re.sub("[^a-zA-Z]", " ", raw_text)

    # convert to lower case and split into words -> convert string into list ( 'hello world' -> ['hello', 'world'])
    words = letters_only_text.lower().split()
    
    cleaned_words = []
    
    # remove stopwords
    for word in words:
        if word not in stop_words:
            cleaned_words.append(word)
    #cleaned_words= []
    #for word in cleaned_word:
        #if word in terms2:
            #cleaned_words.append(word)
    
    # stemm or lemmatise words
    stemmed_words = []
    for w in cleaned_words:
        doc = nlp(w) #dont forget to change stem to lemmatize if you are using a lemmatizer
        for sent in doc.sentences:
            for word in sent.words:
                stemmed_words.append(word.lemma)
    
    # converting list back to string
    return " ".join(stemmed_words)


# In[56]:


test_sentence = "Esta es una oración para demostrar como el preproceso funciona horario vida personal familiar...!"
preprocess(test_sentence)


# In[57]:


df_c['Comentario']=df_c['Comentario'].apply(lambda x: str(x))
df_c['lemma']=df_c['Comentario'].apply(preprocess)
df_c


# In[58]:


from collections import Counter
Counter(" ".join(df_c["lemma"]).split()).most_common(15)


# In[59]:


df_d['dup'] = df_d.duplicated(subset=None, keep='first')
df_d = df_d[df_d['dup'] == False]
df_d= df_d[df_d['Comentario'].str.contains('|'.join(terms1), na=False)]
df_d['Comentario']=df_d['Comentario'].apply(lambda x: str(x))
df_d['lemma']=df_d['Comentario'].apply(preprocess)
from collections import Counter
Counter(" ".join(df_d["lemma"]).split()).most_common(15)


# In[60]:


all_words_c = '' 

#looping through all incidents and joining them to one text, to extract most common words
for arg in df_c["lemma"]: 

    tokens = arg.split()  
      
    all_words_c += " ".join(tokens)+" "


# In[69]:


from nltk.util import ngrams
n_gram = 2
n_gram_dic = dict(Counter(ngrams(all_words_c.split(), n_gram)))

for i in n_gram_dic:
    if n_gram_dic[i] >= 50:
        print(i, n_gram_dic[i])


# In[67]:


all_words_d = '' 

#looping through all incidents and joining them to one text, to extract most common words
for arg in df_d["lemma"]: 

    tokens = arg.split()  
      
    all_words_d += " ".join(tokens)+" "

from nltk.util import ngrams
n_gram = 2
n_gram_dic = dict(Counter(ngrams(all_words_d.split(), n_gram)))

for i in n_gram_dic:
    if n_gram_dic[i] >= 30:
        print(i, n_gram_dic[i])


# In[63]:


pd.set_option("display.max_columns", None)
pd.set_option('display.max_colwidth', None)


# In[64]:


terms3 = ['respetar horario trabajo']
df_d2= df_d[df_d['lemma'].str.contains('|'.join(terms3), na=False)]
df_d2


# In[65]:


terms4 = ['vivir principio cultural']
df_c2= df_c[df_c['lemma'].str.contains('|'.join(terms4), na=False)]
df_c2


# In[70]:


with pd.ExcelWriter('Comentarios.xlsx') as writer1:
    df_c.to_excel(writer1, sheet_name='comenzar', index = False)
    df_d.to_excel(writer1, sheet_name='dejar', index = False)


# In[ ]:




