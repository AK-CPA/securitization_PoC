"""Fuzzy sentence matching for loose language tie-outs.

Extracts sentences from Word documents and finds matches against
candidate sentences using sequence similarity.
"""
import re
from difflib import SequenceMatcher
from docx import Document


def extract_document_text(filepath: str) -> str:
    """Extract all paragraph text from a Word document."""
    doc = Document(filepath)
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def split_into_sentences(text: str) -> list[str]:
    """Split text into sentences using regex.

    Handles common abbreviations and decimal numbers to avoid
    false splits.
    """
    # Protect known abbreviations and decimals
    protected = text
    # Protect common abbreviations
    abbrevs = ["Mr.", "Mrs.", "Ms.", "Dr.", "Inc.", "Ltd.", "Corp.", "vs.", "e.g.", "i.e.", "etc."]
    placeholders = {}
    for i, abbr in enumerate(abbrevs):
        placeholder = f"__ABBR{i}__"
        placeholders[placeholder] = abbr
        protected = protected.replace(abbr, placeholder)

    # Split on sentence-ending punctuation followed by space + capital or end
    raw_sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z"\(])', protected)

    # Restore abbreviations
    sentences = []
    for s in raw_sentences:
        for placeholder, abbr in placeholders.items():
            s = s.replace(placeholder, abbr)
        s = s.strip()
        if s:
            sentences.append(s)

    return sentences


def find_matching_sentences(
    document_text: str,
    candidate_sentences: list[str],
    threshold: float = 0.75,
) -> list[dict]:
    """Find sentences in the document that match the candidate sentences.

    Args:
        document_text: Full text of the Word document
        candidate_sentences: List of reference sentences to match against
        threshold: Minimum similarity ratio (0-1) to consider a match

    Returns:
        List of dicts with:
            - candidate: the original candidate sentence
            - matched_sentence: best matching sentence from document (or None)
            - similarity: float 0-1
            - all_matches: list of (sentence, score) above threshold, sorted by score desc
    """
    doc_sentences = split_into_sentences(document_text)

    results = []
    for candidate in candidate_sentences:
        candidate_clean = _normalize(candidate)

        best_match = None
        best_score = 0.0
        all_matches = []

        for doc_sentence in doc_sentences:
            doc_clean = _normalize(doc_sentence)
            score = SequenceMatcher(None, candidate_clean, doc_clean).ratio()

            if score >= threshold:
                all_matches.append((doc_sentence, score))
                if score > best_score:
                    best_score = score
                    best_match = doc_sentence

        # Sort all matches by score descending
        all_matches.sort(key=lambda x: x[1], reverse=True)

        results.append({
            "candidate": candidate,
            "matched_sentence": best_match,
            "similarity": round(best_score, 4),
            "all_matches": all_matches[:5],  # top 5 matches
        })

    return results


def _normalize(text: str) -> str:
    """Normalize text for comparison: lowercase, collapse whitespace, strip punctuation."""
    text = text.lower()
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    return text
