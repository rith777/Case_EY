import pandas as pd
from collections import defaultdict
import re
import os

def load_data():
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Construct full paths to the Excel files
    nrc_path = os.path.join(script_dir, 'nrc.xlsx')
    tweets_path = os.path.join(script_dir, 'Tweets.xlsx')

    # Load NRC lexicon from Excel
    nrc_df = pd.read_excel(nrc_path)

    # Load tweets data from Excel
    tweets_df = pd.read_excel(tweets_path)

    return nrc_df, tweets_df


def preprocess_text(text):
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r'http\S+|www\S+|https\S+', '', text, flags=re.MULTILINE)
    text = re.sub(r'@\w+', '', text)
    text = re.sub(r'[^\w\s]', '', text)
    return text.strip()


def create_sentiment_dict(nrc_df):
    sentiment_dict = defaultdict(dict)

    # Ensure column names match your actual Excel file
    emotion_columns = ['Positive', 'Negative', 'Anger', 'Anticipation',
                       'Disgust', 'Fear', 'Joy', 'Sadness', 'Surprise', 'Trust']

    for _, row in nrc_df.iterrows():
        word = str(row['English Word']).lower()
        for emotion in emotion_columns:
            # Handle different possible representations of 1/0 (int, float, string)
            emotion_value = row[emotion]
            if pd.notna(emotion_value) and (emotion_value == 1 or str(emotion_value).strip() == '1'):
                sentiment_dict[word][emotion] = 1
    return sentiment_dict


def analyze_tweets(tweets_df, sentiment_dict):
    results = []
    emotion_columns = ['Positive', 'Negative', 'Anger', 'Anticipation',
                       'Disgust', 'Fear', 'Joy', 'Sadness', 'Surprise', 'Trust']

    for _, tweet in tweets_df.iterrows():
        text = preprocess_text(tweet['text'])
        words = text.split()

        emotions = {e: 0 for e in emotion_columns}
        word_count = 0

        for word in words:
            if word in sentiment_dict:
                word_count += 1
                for emotion in sentiment_dict[word]:
                    emotions[emotion] += 1

        # Add normalized scores (per word) as well
        normalized_emotions = {f"{e}_norm": emotions[e] / word_count if word_count > 0 else 0
                               for e in emotion_columns}

        result = {
            'id': tweet['id'],
            'created_at': tweet['created_at'],
            'text': tweet['text'],
            'TwitterName': tweet['TwitterName'],
            'word_count': word_count,
            **emotions,
            **normalized_emotions
        }
        results.append(result)

    return pd.DataFrame(results)


def main():
    try:
        print("Loading data...")
        nrc_df, tweets_df = load_data()

        print("Creating sentiment dictionary...")
        sentiment_dict = create_sentiment_dict(nrc_df)

        print("Analyzing tweets...")
        results_df = analyze_tweets(tweets_df, sentiment_dict)

        # Save to Excel to preserve special characters and formatting
        output_file = 'tweet_sentiment_analysis.xlsx'
        results_df.to_excel(output_file, index=False)
        print(f"Analysis complete. Results saved to {output_file}")

        # Show sample output
        print("\nSample results:")
        print(results_df.head())

    except Exception as e:
        print(f"An error occurred: {str(e)}")


if __name__ == '__main__':
    main()