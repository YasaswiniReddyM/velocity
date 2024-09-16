# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set environment variables to ensure that Python output is sent straight to terminal (without buffering)
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set a working directory for the app
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any needed packages specified in requirements.txt
# Create a requirements.txt containing necessary dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port 5000 for Flask
EXPOSE 5000

# Run the Flask app
CMD ["python", "app.py"]
