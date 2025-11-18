# Use an official Python runtime as a parent image
FROM python:3.10-slim

# Set the working directory in the container
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PORT=9080 \
    HOST=0.0.0.0 \
    FILE_SERVER_PORT=8001 \
    PYTHONUNBUFFERED=1

# Install dependencies
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy the application code
COPY . .

# Create output directory
RUN mkdir -p output

# Expose ports
EXPOSE 9080 8001

# Run the application
CMD ["python", "src/main.py"]
