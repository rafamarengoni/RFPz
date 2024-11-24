import spacy
from transformers import pipeline

# Load spaCy language model
nlp_spacy = spacy.load("en_core_web_sm")


# Load spaCy model
nlp_spacy = spacy.load("en_core_web_sm")

# Load Hugging Face summarization model
summarizer = pipeline("summarization", model="facebook/bart-large-cnn")

def extract_key_details(text):
    """Extract goals, deliverables, timeline, and evaluation criteria using spaCy."""
    doc = nlp_spacy(text)
    goals, deliverables, timeline, criteria = [], [], [], []

    for ent in doc.ents:
        if ent.label_ in ["ORG", "PRODUCT"]:
            goals.append(ent.text)
        elif ent.label_ == "DATE":
            timeline.append(ent.text)
        elif ent.label_ in ["MONEY", "PERCENT"]:
            deliverables.append(ent.text)
        elif "evaluation" in ent.text.lower():
            criteria.append(ent.text)

    return {
        "Goals": " | ".join(goals) if goals else "No specific goals detected.",
        "Deliverables": " | ".join(deliverables) if deliverables else "No specific deliverables detected.",
        "Timeline": " | ".join(timeline) if timeline else "No timeline detected.",
        "Evaluation Criteria": " | ".join(criteria) if criteria else "No evaluation criteria detected.",
    }

def summarize_text(text, max_length=130, min_length=30):
    """Summarize long text using a Hugging Face model."""
    if len(text.split()) < 50:  # Skip summarization for very short text
        return text

    summary = summarizer(
        text,
        max_length=max_length,
        min_length=min_length,
        do_sample=False,
    )
    return summary[0]["summary_text"]
