import os
import sys
from docx import Document
from transformers import pipeline

# Initialize the summarization pipeline once at module level
try:
    summarizer = pipeline("summarization", model="facebook/bart-large-cnn")
except Exception as e:
    print(f"Error initializing summarization model: {e}")
    print("Please ensure you have internet connection for model download.")
    sys.exit(1)

def load_document(file_path):
    """
    Load a DOCX document.
    Returns the text content of the document.
    Note: Only supports standard DOCX format, not proprietary WPS formats.
    """
    try:
        doc = Document(file_path)
        return '\n'.join([para.text for para in doc.paragraphs if para.text])
    except Exception as e:
        print(f"Error loading document: {e}")
        print("Note: This tool only supports standard DOCX files, not proprietary WPS formats.")
        sys.exit(1)

def generate_summary(text, max_length=150):
    """
    Generate a summary of the input text using a transformer model.
    Returns a summarized version of the text.
    """
    try:
        if not text or len(text.strip()) < 30:
            print("Text too short for summarization")
            return None
            
        # Truncate text if too long (typical model limit ~1024 tokens)
        if len(text) > 4000:  # rough character limit
            text = text[:4000]
            
        summary = summarizer(text, max_length=max_length, min_length=30, do_sample=False)
        return summary[0]['summary_text']
    except Exception as e:
        print(f"Error during summarization: {e}")
        return None

def main(file_path):
    """
    Main function to handle the flow of the summary generation.
    """
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        sys.exit(1)

    print("Loading document...")
    document_text = load_document(file_path)

    if not document_text:
        print("Document is empty or could not be read.")
        sys.exit(1)

    print("Generating summary...")
    summary = generate_summary(document_text)

    if summary:
        print("Summary generated successfully:\n")
        print(summary)
    else:
        print("Failed to generate summary.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python summary_generator.py <path_to_docx>")
        sys.exit(1)

    main(sys.argv[1])
