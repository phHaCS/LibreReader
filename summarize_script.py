import sys
from transformers import pipeline

def summarize_document(document_path):
    with open(document_path, 'r', encoding='utf-8', errors='ignore') as file:
        content = file.read()

    summarizer = pipeline('summarization')
    summary = summarizer(content, max_length=130, min_length=30, do_sample=False)

    return summary[0]['summary_text']

if __name__ == "__main__":
    document_path = sys.argv[1]  # Get the document path from the command-line arguments
    summary_text = summarize_document(document_path)
    print(summary_text)