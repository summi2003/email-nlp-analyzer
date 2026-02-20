# Email NLP Analyzer

## Overview
This project is a Natural Language Processing (NLP) based system that connects to an email inbox, processes email text, and extracts important keywords using TF-IDF.

The goal is to convert unstructured email communication into structured analytical insights.
## Machine Learning Concept Used
This project uses TF-IDF (Term Frequency–Inverse Document Frequency) to identify important words in email communication.  
TF-IDF assigns higher weight to words that are frequent in a specific email but rare across other emails, helping detect meaningful topics.

## Features
- Fetches emails securely using IMAP
- Cleans and preprocesses email text
- Extracts top keywords using TF-IDF
- Identifies keyword occurrence (subject/body)
- Generates Excel analysis reports

## Tech Stack
- Python
- NLP Text Processing
- TF-IDF Vectorization
- Pandas
- Scikit-learn

## How to Run
1. Install dependencies
