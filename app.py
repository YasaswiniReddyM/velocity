from flask import Flask, render_template, request, send_file
import pandas as pd

app = Flask(__name__)

def parse_dates(date, formats):
    for fmt in formats:
        try:
            return pd.to_datetime(date, format=fmt)
        except ValueError:
            continue
    return pd.NaT  # Return Not a Time if all formats fail

formats = ['%d-%m-%Y %H:%M:%S', '%m/%d/%Y %H:%M', '%Y-%m-%d']  # Add formats as needed

def process_tickets_and_incidents(problem_tickets_file, incidents_file):
    # Load the problem ticket CSV file
    problem_tickets_df = pd.read_csv(problem_tickets_file)

    # Load the incident CSV file
    incidents_df = pd.read_csv(incidents_file)

    # Convert date columns to datetime type
    problem_tickets_df['Opened'] = pd.to_datetime(problem_tickets_df['Opened'], errors='coerce')
    problem_tickets_df['Closed'] = pd.to_datetime(problem_tickets_df['Closed'], errors='coerce')
    incidents_df['Opened'] = pd.to_datetime(incidents_df['Opened'], errors='coerce')

    # Initialize a dictionary to store the count of incidents for each problem ticket
    problem_ticket_incident_count = {}

    # Iterate over each problem ticket
    for index, ticket_row in problem_tickets_df.iterrows():
        # Filter incidents where tag matches problem ticket tag and opened date is between ticket's opened and closed dates
        matching_incidents = incidents_df[(incidents_df['Tags'] == ticket_row['Tags']) &
                                          (incidents_df['Opened'] >= ticket_row['Opened']) &
                                          (incidents_df['Opened'] <= ticket_row['Closed'])]

        # Count the number of matching incidents
        incident_count = len(matching_incidents)

        # Store the count in the dictionary
        problem_ticket_incident_count[ticket_row['Number']] = incident_count

    # Convert the dictionary to a DataFrame
    incident_count_df = pd.DataFrame.from_dict(problem_ticket_incident_count, orient='index',
                                               columns=['incident_count'])

    # Merge the incident count DataFrame with the problem ticket DataFrame
    merged_df = pd.merge(problem_tickets_df, incident_count_df, left_on='Number', right_index=True, how='left')

    # Save the DataFrame to a CSV file
    output_csv_file = 'output.csv'
    merged_df.to_csv(output_csv_file, index=False)  # Set index=False to avoid saving DataFrame index as a separate column in the CSV

    return output_csv_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        problem_csv = request.files['problem_csv']
        incidents_csv = request.files['incidents_csv']

        if not (problem_csv and incidents_csv):
            return "Files are required.", 400

        problem_csv.save('problem.csv')
        incidents_csv.save('incidents.csv')

        output_file = process_tickets_and_incidents('problem.csv', 'incidents.csv')
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)

