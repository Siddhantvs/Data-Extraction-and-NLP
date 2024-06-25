from os import path
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

from nltk.tokenize import *


stopwords_file = open("stopwords.txt", "r")
stopwords = stopwords_file.read().splitlines()
stopwords = [word.lower() for word in stopwords]

stopwords_set = set(stopwords)

positivewords_file = open("positivewords.txt", "r")
positivewords = positivewords_file.read().splitlines()
positivewords = [word.lower() for word in positivewords]

positivewords_set = set(positivewords)

negativewords_file = open("negativewords.txt", "r")
negativewords = negativewords_file.read().splitlines()
negativewords = [word.lower() for word in negativewords]

negativewords_set = set(negativewords)

inputxlsx = pd.read_excel(r"C:\S\blackcofferproj\Input.xlsx")

results_data = {
    'URL_ID': [],
    'URL': [],
    'POSITIVE SCORE': [],
    'NEGATIVE SCORE': [],
    'POLARITY SCORE': [],
    'SUBJECTIVITY SCORE': [],
    'AVG SENTENCE LENGTH': [],
    'PERCENTAGE OF COMPLEX WORDS': [],
    'FOG INDEX': [],
    'AVG NUMBER OF WORDS PER SENTENCE': [],
    'COMPLEX WORD COUNT': [],
    'WORD COUNT': [],
    'SYLLABLE PER WORD': [],
    'PERSONAL PRONOUNS': [],
    'AVG WORD LENGTH': []
}


for index, row in inputxlsx.iterrows():
    print(row["URL_ID"], row["URL"])
    url_link = row["URL"]
    try:
        result = requests.get(url_link)
        result.encoding = 'utf-8'
    except:
        print("404")

    else:
        file_path = path.relpath(
            "titlecontents/" + row["URL_ID"] + ".txt"
        )

        result_content = result.text

        if len(result_content) == 0:
            results_data['URL_ID'].append(row['URL_ID'])
            results_data['URL'].append(row['URL'])
            results_data['POSITIVE SCORE'].append('Page Empty')
            results_data['NEGATIVE SCORE'].append('Page Empty')
            results_data['POLARITY SCORE'].append('Page Empty')
            results_data['SUBJECTIVITY SCORE'].append('Page Empty')
            results_data['AVG SENTENCE LENGTH'].append('Page Empty')
            results_data['PERCENTAGE OF COMPLEX WORDS'].append('Page Empty')
            results_data['FOG INDEX'].append('Page Empty')
            results_data['AVG NUMBER OF WORDS PER SENTENCE'].append('Page Empty')
            results_data['COMPLEX WORD COUNT'].append('Page Empty')
            results_data['WORD COUNT'].append('Page Empty')
            results_data['SYLLABLE PER WORD'].append('Page Empty')
            results_data['PERSONAL PRONOUNS'].append('Page Empty')
            results_data['AVG WORD LENGTH'].append('Page Empty')
            continue
        soup = BeautifulSoup(result_content, "html.parser")

        soup.prettify(formatter=lambda s: s.replace(u'\xa0', ' '))

        f = open(file_path, "w", errors="ignore")

        # traverse paragraphs from soup
        for data in soup.find_all("h1", {"class": "entry-title"}):
            sum = data.get_text()
            f.writelines(sum)

        for data in soup.find_all("div", {"class": "td-post-content tagdiv-type"}):
            sum = data.get_text()
            f.writelines(sum)

        f.close()

        fnew = open(file_path, "r", errors="ignore")
        extracted_text = fnew.read()

        words = word_tokenize(extracted_text)
        sentences = sent_tokenize(extracted_text)

        sentences_count = len(sentences)

        original_words = words.copy()
        words = [word.lower() for word in words]

        cleaned_words = [word for word in words if word not in stopwords_set]

        nonPunct = re.compile('.*[A-Za-z0-9].*')
        cleaned_words_nonPunct = [
            w for w in cleaned_words if nonPunct.match(w)]
        words_nonPunct = [w for w in words if nonPunct.match(w)]

        words_count = len(words_nonPunct)

        if (words_count) <= 0:
            results_data['URL_ID'].append(row['URL_ID'])
            results_data['URL'].append(row['URL'])
            results_data['POSITIVE SCORE'].append('Page Empty')
            results_data['NEGATIVE SCORE'].append('Page Empty')
            results_data['POLARITY SCORE'].append('Page Empty')
            results_data['SUBJECTIVITY SCORE'].append('Page Empty')
            results_data['AVG SENTENCE LENGTH'].append('Page Empty')
            results_data['PERCENTAGE OF COMPLEX WORDS'].append('Page Empty')
            results_data['FOG INDEX'].append('Page Empty')
            results_data['AVG NUMBER OF WORDS PER SENTENCE'].append('Page Empty')
            results_data['COMPLEX WORD COUNT'].append('Page Empty')
            results_data['WORD COUNT'].append('Page Empty')
            results_data['SYLLABLE PER WORD'].append('Page Empty')
            results_data['PERSONAL PRONOUNS'].append('Page Empty')
            results_data['AVG WORD LENGTH'].append('Page Empty')
            continue

        cleaned_words_count = len(cleaned_words_nonPunct)

        positive_score = 0
        negative_score = 0
        for word in cleaned_words_nonPunct:
            if word in positivewords_set:
                positive_score += 1
            if word in negativewords_set:
                negative_score += 1

        polarity_score = -1
        if positive_score + negative_score != 0:
            polarity_score = ((positive_score - negative_score) /
                              (positive_score + negative_score)) + 0.000001

        subjectivity_score = (
            (positive_score + negative_score) / cleaned_words_count) + 0.000001

        avg_sentence_length = words_count / sentences_count

        complex_word_count = 0
        total_syllable_count = 0
        for word in words_nonPunct:
            vowel_count = 0
            vowels = set(['a', 'e', 'i', 'o', 'u'])
            for c in word:
                if c in vowels:
                    vowel_count += 1
            if len(word) > 2:
                if (word[-1] == 'd' or word[-1] == 's') and word[-2] == 'e':
                    vowel_count -= 1

            total_syllable_count += vowel_count
            if (vowel_count > 2):
                complex_word_count += 1

        percentage_of_complex_words = complex_word_count / words_count
        fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)

        syllable_count_per_word = total_syllable_count / words_count

        personal_pronouns = ['I', 'we', 'ours', 'my', 'us']
        personal_pronoun_count = 0
        for word in original_words:
            if word in personal_pronouns:
                personal_pronoun_count += 1

        total_char_count = 0
        for word in words_nonPunct:
            total_char_count += len(word)

        avg_word_length = total_char_count / words_count
   

        results_data['URL_ID'].append(row['URL_ID'])
        results_data['URL'].append(row['URL'])
        results_data['POSITIVE SCORE'].append(positive_score)
        results_data['NEGATIVE SCORE'].append(negative_score)
        results_data['POLARITY SCORE'].append(polarity_score)
        results_data['SUBJECTIVITY SCORE'].append(subjectivity_score)
        results_data['AVG SENTENCE LENGTH'].append(avg_sentence_length)
        results_data['PERCENTAGE OF COMPLEX WORDS'].append(percentage_of_complex_words)
        results_data['FOG INDEX'].append(fog_index)
        results_data['AVG NUMBER OF WORDS PER SENTENCE'].append(avg_sentence_length)
        results_data['COMPLEX WORD COUNT'].append(complex_word_count)
        results_data['WORD COUNT'].append(cleaned_words_count)
        results_data['SYLLABLE PER WORD'].append(syllable_count_per_word)
        results_data['PERSONAL PRONOUNS'].append(personal_pronoun_count)
        results_data['AVG WORD LENGTH'].append(avg_word_length)

results_df = pd.DataFrame(results_data)

results_df.to_excel(r'C:\S\blackcofferproj\Output.xlsx', sheet_name='sheet1', index=False)