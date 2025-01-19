# Use the official Python base image
FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libffi-dev \
    libgmp-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy the application code to the container
COPY . /app

# Install Python dependencies
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Set environment variables
ENV FLASK_APP=app.py
ENV FLASK_ENV=production

# Expose the port for the Flask app
EXPOSE 5000

# Install Gunicorn (a production WSGI server for Flask)
RUN pip install gunicorn

# Run the Flask app using Gunicorn, bind to all interfaces
CMD ["gunicorn", "-b", "0.0.0.0:5000", "app:app"]
