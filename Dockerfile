# # Use an official Python runtime as base
# FROM python:3.10-slim

# # Set environment variables
# ENV PYTHONDONTWRITEBYTECODE=1
# ENV PYTHONUNBUFFERED=1
# ENV STREAMLIT_SERVER_HEADLESS=true
# ENV PORT=8080

# # Set working directory
# WORKDIR /app

# # Install system dependencies
# RUN apt-get update && apt-get install -y \
#     libreoffice \
#     python3-pip \
#     build-essential \
#     libgl1-mesa-glx \
#     libglib2.0-0 \
#     # Additional dependencies for PyMuPDF
#     swig \
#     libfreetype6-dev \
#     libharfbuzz-dev \
#     libfribidi-dev \
#     libmupdf-dev \
#     && apt-get clean

# # Copy only requirements to leverage Docker cache
# COPY requirements.txt .

# # Install Python dependencies
# RUN pip install --upgrade pip
# RUN pip install -r requirements.txt

# # Copy the rest of the application
# COPY . .

# # Expose port 8080 for Cloud Run
# EXPOSE 8080

# # Run Streamlit app on port 8080
# CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.enableCORS=false", "--server.enableXsrfProtection=false"]

# Use an official Python runtime as base
FROM python:3.10-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV STREAMLIT_SERVER_HEADLESS=true
ENV PORT=8080

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    python3-pip \
    build-essential \
    libgl1-mesa-glx \
    libglib2.0-0 \
    swig \
    libfreetype6-dev \
    libharfbuzz-dev \
    libfribidi-dev \
    libmupdf-dev \
    && ln -s /usr/bin/libreoffice /usr/bin/soffice \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy app code
COPY . .

# Expose port for Cloud Run
EXPOSE 8080

# Run the app
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.enableCORS=false", "--server.enableXsrfProtection=false"]
