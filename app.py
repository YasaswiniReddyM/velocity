from flask import Flask, render_template, request, send_file, send_from_directory
import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

app = Flask(__name__)

# Load the pre-trained model from Sentence Transformers
model = SentenceTransformer('distilbert-base-nli-mean-tokens')

def process_tickets_and_incidents(problem_tickets_file, incidents_file):
    # Load the problem ticket Excel file
    problem_tickets_df = pd.read_excel(problem_tickets_file)

    # Load the incident Excel file
    incidents_df = pd.read_excel(incidents_file)

    # Prepare text data from problem tickets and incidents
    problem_ticket_texts = (problem_tickets_df['Problem statement'].fillna('') + " " + problem_tickets_df['Tags'].fillna('')).tolist()

    short_description_texts = incidents_df['Short description'].fillna('').tolist()
    description_texts = incidents_df['Description'].fillna('').tolist()
    resolution_notes_texts = incidents_df['Resolution notes'].fillna('').tolist()
    tag_texts = incidents_df['Tags'].fillna('').tolist()

    # Get embeddings for problem ticket texts
    problem_ticket_embeddings = model.encode(problem_ticket_texts)

    # Get embeddings for incident texts
    short_description_embeddings = model.encode(short_description_texts)
    description_embeddings = model.encode(description_texts)
    resolution_notes_embeddings = model.encode(resolution_notes_texts)
    tag_embeddings = model.encode(tag_texts)

    # Compute cosine similarity
    similarity_short_description = cosine_similarity(problem_ticket_embeddings, short_description_embeddings)
    similarity_description = cosine_similarity(problem_ticket_embeddings, description_embeddings)
    similarity_resolution_notes = cosine_similarity(problem_ticket_embeddings, resolution_notes_embeddings)
    similarity_tag = cosine_similarity(problem_ticket_embeddings, tag_embeddings)

    # Define similarity threshold
    threshold = 0.87

    # Count number of incidents impacted and list incident numbers for each problem ticket
    num_impacted_incidents = []
    impacted_incident_numbers = []

    for i in range(len(problem_ticket_texts)):
        impacted_count = 0
        impacted_numbers = []

        for j in range(len(short_description_texts)):
            # Check if any similarity score exceeds the threshold
            if (similarity_short_description[i, j] >= threshold or
                similarity_description[i, j] >= threshold or
                similarity_resolution_notes[i, j] >= threshold or
                similarity_tag[i, j] >= threshold):
                
                impacted_count += 1
                impacted_numbers.append(incidents_df.iloc[j]['Number'])  # Assuming 'Number' is the incident ID

        num_impacted_incidents.append(impacted_count)
        impacted_incident_numbers.append(impacted_numbers)

    # Add the results to the problem_tickets_df
    problem_tickets_df['num_impacted_incidents'] = num_impacted_incidents
    problem_tickets_df['impacted_incident_numbers'] = impacted_incident_numbers

    # Save the DataFrame to an Excel file
    output_excel_file = 'output.xlsx'
    problem_tickets_df.to_excel(output_excel_file, index=False)

    return output_excel_file

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        problem_excel = request.files['problem_excel']
        incidents_excel = request.files['incidents_excel']

        if not (problem_excel and incidents_excel):
            return "Files are required.", 400

        problem_excel.save('problem.xlsx')
        incidents_excel.save('incidents.xlsx')

        output_file = process_tickets_and_incidents('problem.xlsx', 'incidents.xlsx')
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
