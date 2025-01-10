import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import re
from textblob import TextBlob

# Helper functions for textual analysis
def count_syllables(word):
    word = word.lower()
    vowels = "aeiouy"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count = 1
    return count

def is_complex_word(word):
    return count_syllables(word) > 2

def calculate_fog_index(avg_sentence_length, percentage_complex_words):
    return 0.4 * (avg_sentence_length + percentage_complex_words)

def count_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    return len(pronouns)

# Load the input Excel file
input_file_path = 'Input.xlsx'  # Update with your uploaded file path
output_directory = 'articles'  # Directory to save articles
output_data_path = 'output_analysis.xlsx'  # Output Excel file for analysis results

# Create output directory if it does not exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Read the Excel file
urls_df = pd.read_excel(input_file_path)

# Data structure for storing analysis
analysis_data = []

# Check if the required columns exist
if 'URL_ID' not in urls_df.columns or 'URL' not in urls_df.columns:
    print("Input file must contain 'URL_ID' and 'URL' columns.")
else:
    # Iterate through each row and extract article text
    for index, row in urls_df.iterrows():
        url_id = str(row['URL_ID'])
        url = row['URL']

        try:
            # Fetch the content of the URL
            response = requests.get(url)
            response.raise_for_status()  # Raise an error for bad responses

            # Parse the HTML content using BeautifulSoup
            soup = BeautifulSoup(response.content, 'html.parser')

            # Extract the title and article body
            title = soup.find('h1').get_text(strip=True) if soup.find('h1') else 'No Title Found'
            article_body = '\n'.join([p.get_text(strip=True) for p in soup.find_all('p')])

            if not article_body.strip():
                article_body = 'No Article Text Found'

            # Combine title and body
            article_text = f"{title}\n\n{article_body}"

            # Save the article text in a text file named by URL_ID
            output_file_path = os.path.join(output_directory, f"{url_id}.txt")
            with open(output_file_path, 'w', encoding='utf-8') as file:
                file.write(article_text)

            # Perform textual analysis
            words = re.findall(r'\w+', article_text)
            word_count = len(words)
            char_count = len(article_text)
            sentence_count = len(re.split(r'[.!?]', article_text)) - 1
            avg_sentence_length = word_count / sentence_count if sentence_count > 0 else 0
            complex_word_count = sum(1 for word in words if is_complex_word(word))
            percentage_complex_words = (complex_word_count / word_count) * 100 if word_count > 0 else 0
            fog_index = calculate_fog_index(avg_sentence_length, percentage_complex_words)
            avg_word_length = char_count / word_count if word_count > 0 else 0
            syllables_per_word = sum(count_syllables(word) for word in words) / word_count if word_count > 0 else 0
            personal_pronouns = count_personal_pronouns(article_text)
            sentiment = TextBlob(article_text).sentiment
            polarity_score = sentiment.polarity
            subjectivity_score = sentiment.subjectivity

            # Save analysis results
            analysis_data.append({
                'URL_ID': url_id,
                'Title': title,
                'Positive Score': sum(1 for word in words if TextBlob(word).sentiment.polarity > 0),
                'Negative Score': sum(1 for word in words if TextBlob(word).sentiment.polarity < 0),
                'Polarity Score': polarity_score,
                'Subjectivity Score': subjectivity_score,
                'Average Sentence Length': avg_sentence_length,
                'Percentage of Complex Words': percentage_complex_words,
                'Fog Index': fog_index,
                'Average Number of Words per Sentence': avg_sentence_length,
                'Complex Word Count': complex_word_count,
                'Word Count': word_count,
                'Syllables per Word': syllables_per_word,
                'Personal Pronouns': personal_pronouns,
                'Average Word Length': avg_word_length
            })

            print(f"Successfully saved and analyzed article {url_id}.txt")

        except Exception as e:
            print(f"Failed to process URL ID {url_id}: {e}")

# Save the analysis data to an Excel file
analysis_df = pd.DataFrame(analysis_data)
analysis_df.to_excel(output_data_path, index=False)

print(f"Articles saved in {output_directory}")
print(f"Textual analysis saved in {output_data_path}")
