from flask import Flask, render_template, request, send_file
import pandas as pd
import re
from collections import Counter
import nltk
from nltk.corpus import wordnet
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

app = Flask(__name__)

nltk.download('wordnet')

def expand_synonyms(text):
    words = re.findall(r'\w+', text.lower())
    expanded_words = set(words)
    for word in words:
        synonyms = wordnet.synsets(word)
        for syn in synonyms:
            for lemma in syn.lemmas():
                expanded_words.add(lemma.name())  # Add synonyms to the set
    return expanded_words

def process_tickets_and_incidents(problem_tickets_file, incidents_file):
    # Load the Excel files
    problem_tickets_df = pd.read_excel(problem_tickets_file)
    incidents_df = pd.read_excel(incidents_file)

    # Combine problem statement and tags for problem tickets
    problem_ticket_texts = (problem_tickets_df['Problem statement'].fillna('') + " " + problem_tickets_df['Tags'].fillna('')).tolist()

    # Combine short description, description, resolution notes, and tags for incidents
    incident_texts = (incidents_df['Short description'].fillna('') + " " + 
                      incidents_df['Description'].fillna('') + " " +
                      incidents_df['Resolution notes'].fillna('') + " " +
                      incidents_df['Tags'].fillna('')).tolist()

    # Create a TF-IDF vectorizer
    vectorizer = TfidfVectorizer(ngram_range=(1, 3))

    # Fit the vectorizer on both problem tickets and incidents
    combined_texts = problem_ticket_texts + incident_texts
    tfidf_matrix = vectorizer.fit_transform(combined_texts)

    # Split the matrix back into problem ticket vectors and incident vectors
    problem_ticket_tfidf = tfidf_matrix[:len(problem_ticket_texts)]
    incident_tfidf = tfidf_matrix[len(problem_ticket_texts):]

    # Compute cosine similarity between problem tickets and incidents
    similarity_matrix = cosine_similarity(problem_ticket_tfidf, incident_tfidf)

    # Define a similarity threshold
    threshold = 0.4  # Lowered threshold due to using TF-IDF

    # Count impacted incidents and list incident numbers for each problem ticket
    num_impacted_incidents = []
    impacted_incident_numbers = []

    for i in range(len(problem_ticket_texts)):
        impacted_count = 0
        impacted_numbers = []

        for j in range(len(incident_texts)):
            if similarity_matrix[i, j] >= threshold:
                impacted_count += 1
                impacted_numbers.append(incidents_df.iloc[j]['Number'])

        num_impacted_incidents.append(impacted_count)
        impacted_incident_numbers.append(impacted_numbers)

    # Add the results to the problem_tickets_df
    problem_tickets_df['num_impacted_incidents'] = num_impacted_incidents
    problem_tickets_df['impacted_incident_numbers'] = impacted_incident_numbers

    # Save the DataFrame to an Excel file
    output_excel_file = 'output_tfidf.xlsx'
    problem_tickets_df.to_excel(output_excel_file, index=False)

    return output_excel_file

# def preprocess_text(text):
#     # Convert to lowercase and split into words
#     words = re.findall(r'\w+', text.lower())
#     return set(words)  # Use set for faster lookups

def jaccard_similarity(set1, set2):
    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))
    return intersection / union if union != 0 else 0

# def process_tickets_and_incidents(problem_tickets_file, incidents_file):
#     # Load the Excel files
#     problem_tickets_df = pd.read_excel(problem_tickets_file)
#     incidents_df = pd.read_excel(incidents_file)

#     # Preprocess problem ticket texts
#     problem_ticket_texts = (problem_tickets_df['Problem statement'].fillna('') + " " + problem_tickets_df['Tags'].fillna('')).apply(preprocess_text)

#     # Preprocess incident texts
#     incident_texts = (
#         incidents_df['Short description'].fillna('') + " " +
#         incidents_df['Description'].fillna('') + " " +
#         incidents_df['Resolution notes'].fillna('') + " " +
#         incidents_df['Tags'].fillna('')
#     ).apply(preprocess_text)

#     # Define similarity threshold
#     threshold = 0.2  # Adjust this value based on your needs

#     # Count number of incidents impacted and list incident numbers for each problem ticket
#     results = []
#     for problem_text in problem_ticket_texts:
#         impacted_incidents = []
#         for idx, incident_text in enumerate(incident_texts):
#             similarity = jaccard_similarity(problem_text, incident_text)
#             if similarity >= threshold:
#                 impacted_incidents.append(incidents_df.iloc[idx]['Number'])
        
#         results.append({
#             'num_impacted_incidents': len(impacted_incidents),
#             'impacted_incident_numbers': impacted_incidents
#         })

#     # Add the results to the problem_tickets_df
#     problem_tickets_df['num_impacted_incidents'] = [r['num_impacted_incidents'] for r in results]
#     problem_tickets_df['impacted_incident_numbers'] = [r['impacted_incident_numbers'] for r in results]

#     # Save the DataFrame to an Excel file
#     output_excel_file = 'output.xlsx'
#     problem_tickets_df.to_excel(output_excel_file, index=False)

#     return output_excel_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        problem_excel = request.files['problem_excel']
        incidents_excel = request.files['incidents_excel']

        if not (problem_excel and incidents_excel):
            return "Files are required.", 400

        problem_excel.save('problem.xlsx')
        incidents_excel.save('incidents.excel')

        output_file = process_tickets_and_incidents('problem.xlsx', 'incidents.excel')
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
