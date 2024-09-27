from flask import Flask, render_template, request, send_file
import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# Initialize the Flask app
app = Flask(__name__)

# Load the pre-trained model from Sentence Transformers
model = SentenceTransformer('distilbert-base-nli-mean-tokens')

def process_tickets_and_incidents(problem_tickets_file, incidents_file):
    # Load the problem ticket Excel file (with openpyxl engine specified)
    problem_tickets_df = pd.read_excel(problem_tickets_file, engine='openpyxl')

    # Load the incident Excel file (with openpyxl engine specified)
    incidents_df = pd.read_excel(incidents_file, engine='openpyxl')

    # Prepare text data
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
            if (similarity_short_description[i, j] >= threshold or
                similarity_description[i, j] >= threshold or
                similarity_resolution_notes[i, j] >= threshold or
                similarity_tag[i, j] >= threshold):
                impacted_count += 1
                impacted_numbers.append(incidents_df.iloc[j]['Number'])

        num_impacted_incidents.append(impacted_count)
        impacted_incident_numbers.append(impacted_numbers)

    # Add the results to the problem_tickets_df
    problem_tickets_df['num_impacted_incidents'] = num_impacted_incidents
    problem_tickets_df['impacted_incident_numbers'] = impacted_incident_numbers

    # Save the DataFrame to an Excel file (with openpyxl engine specified)
    output_excel_file = 'output.xlsx'
    problem_tickets_df.to_excel(output_excel_file, engine='openpyxl', index=False)

    return output_excel_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # Get the uploaded Excel files
        problem_excel = request.files['problem_excel']
        incidents_excel = request.files['incidents_excel']

        # Ensure both files are provided
        if not (problem_excel and incidents_excel):
            return "Both Problem Ticket and Incident files are required.", 400

        # Save the uploaded files to disk
        problem_excel.save('problem.xlsx')
        incidents_excel.save('incidents.xlsx')

        # Process the files
        output_file = process_tickets_and_incidents('problem.xlsx', 'incidents.xlsx')

        # Send the processed file back to the user
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
