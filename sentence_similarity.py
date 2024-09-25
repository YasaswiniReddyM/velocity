import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# Load the pre-trained model from Sentence Transformers
model = SentenceTransformer('distilbert-base-nli-mean-tokens')

# Load Excel files
problem_tickets_df = pd.read_excel("path/to/problem.xlsx")  # Update the path
incidents_df = pd.read_excel("path/to/incident.xlsx")  # Update the path

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

# Output the result
print(problem_tickets_df[['Number', 'num_impacted_incidents', 'impacted_incident_numbers']])

# Save the DataFrame to an Excel file
excel_file_name = 'sample_data.xlsx'
problem_tickets_df.to_excel(excel_file_name, index=False)  # Set index=False to avoid saving DataFrame index as a separate column in the Excel file
