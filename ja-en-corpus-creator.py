import os
import re
import argparse
import pdfplumber
import pandas as pd
import spacy
import numpy as np
from tqdm import tqdm
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_corpus_creator.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("pdf_corpus_creator")

def load_spacy_models():
    """Load and return spaCy language models for Japanese and English."""
    models = {}
    
    # Load Japanese model
    try:
        models['ja'] = spacy.load('ja_core_news_sm')
        logger.info("Loaded Japanese language model")
    except OSError:
        logger.error("Japanese model not found. Please install it with: python -m spacy download ja_core_news_sm")
        raise
    
    # Load English model
    try:
        models['en'] = spacy.load('en_core_web_sm')
        logger.info("Loaded English language model")
    except OSError:
        logger.error("English model not found. Please install it with: python -m spacy download en_core_web_sm")
        raise
    
    return models

def extract_text_from_pdf(pdf_path, detect_columns=True, detect_tables=True):
    """Extract text from a PDF file with options for handling columns and tables."""
    logger.info(f"Extracting text from: {pdf_path}")
    all_text = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(tqdm(pdf.pages, desc="Reading pages")):
                page_text = ""
                
                if detect_tables:
                    # Extract tables first
                    tables = page.extract_tables()
                    for table in tables:
                        # Convert table to text representation
                        table_text = "\n".join([" | ".join([str(cell) if cell else "" for cell in row]) for row in table])
                        page_text += table_text + "\n\n"
                
                # Extract remaining text
                if detect_columns:
                    # Try to detect columns using horizontal spaces
                    words = page.extract_words(keep_blank_chars=True)
                    if words:
                        # Sort words by y-position (top to bottom)
                        words_by_line = {}
                        for word in words:
                            y1 = round(word['top'], 0)
                            if y1 not in words_by_line:
                                words_by_line[y1] = []
                            words_by_line[y1].append(word)
                        
                        # Process each line
                        for y, line_words in sorted(words_by_line.items()):
                            # Sort words by x-position (left to right)
                            sorted_words = sorted(line_words, key=lambda w: w['x0'])
                            line_text = " ".join(w['text'] for w in sorted_words)
                            page_text += line_text + "\n"
                else:
                    # Standard text extraction
                    extracted = page.extract_text()
                    if extracted:
                        page_text += extracted
                
                all_text.append(page_text)
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        return None
    
    full_text = "\n\n".join(all_text).strip()
    logger.info(f"Extracted {len(full_text)} characters from {pdf_path}")
    return full_text

def clean_text(text, language):
    """Clean the extracted text."""
    if not text:
        return text
    
    logger.info(f"Cleaning {language} text...")
    
    # Remove excess whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Remove page numbers, headers, footers
    text = re.sub(r'\d+\s*$', '', text)  # Page numbers at the end of lines
    text = re.sub(r'^\s*\d+\s*', '', text)  # Page numbers at the beginning of lines
    text = re.sub(r'Copyright © \d{4}.*', '', text)
    text = re.sub(r'Page \d+ of \d+', '', text)
    
    # Language-specific cleaning
    if language.lower() == 'ja':
        # Remove unnecessary spaces in Japanese text
        text = re.sub(r'([^\w\s])\s+', r'\1', text)
        text = re.sub(r'\s+([^\w\s])', r'\1', text)
    
    return text

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)$', 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)$', 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part)\s+\d+\s*:?\s*(.+)$', 'labeled'),  # Labeled headings
    ]
    
    in_table = False
    table_content = []
    
    for line in lines:
        line = line.strip()
        if not line:
            if in_table and table_content:
                # End of table
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row (contains multiple | characters)
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                in_table = True
            table_content.append(line)
            continue
        elif in_table:
            # End of table
            structured_elements.append({
                'text': '\n'.join(table_content),
                'type': 'table',
                'level': 0
            })
            table_content = []
            in_table = False
        
        # Check for headings
        is_heading = False
        heading_level = 0
        
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                is_heading = True
                if h_type == 'markdown':
                    # Count # symbols for heading level
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    # Count dots for level
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1  # Default level
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if len(line) < 100 and (line.isupper() or line.istitle()):
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Regular paragraph text
                if structured_elements and structured_elements[-1]['type'] == 'paragraph':
                    # Append to previous paragraph
                    structured_elements[-1]['text'] += ' ' + line
                else:
                    structured_elements.append({
                        'text': line,
                        'type': 'paragraph',
                        'level': 0
                    })
    
    # Handle any remaining table content
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    logger.info(f"Detected {len(structured_elements)} structural elements in {language} text")
    return structured_elements

def segment_into_sentences(text, language, spacy_models):
    """Split text into sentences using spaCy."""
    if not text:
        return []
    
    try:
        if language.lower() == 'ja':
            doc = spacy_models['ja'](text)
        else:  # Assume English
            doc = spacy_models['en'](text)
        
        # Extract sentences
        sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
        return sentences
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {e}")
        logger.info("Falling back to rule-based sentence splitting")
        
        # Fallback to simple sentence splitting
        if language.lower() == 'ja':
            # Japanese sentence end markers
            split_pattern = r'([。！？])'
        else:
            # English sentence end markers
            split_pattern = r'([.!?])\s'
        
        sentences = re.split(split_pattern, text)
        # Recombine sentence with its punctuation
        result = []
        i = 0
        while i < len(sentences) - 1:
            if i + 1 < len(sentences) and re.match(r'[.!?。！？]', sentences[i+1]):
                result.append(sentences[i] + sentences[i+1])
                i += 2
            else:
                result.append(sentences[i])
                i += 1
        
        if i < len(sentences):
            result.append(sentences[i])
        
        return [s.strip() for s in result if s.strip()]

def align_sentences_dynamic(ja_sentences, en_sentences, similarity_threshold=0.3):
    """
    Align Japanese and English sentences using dynamic programming and similarity metrics.
    """
    if not ja_sentences or not en_sentences:
        return []
    
    logger.info(f"Aligning {len(ja_sentences)} Japanese sentences with {len(en_sentences)} English sentences")
    
    # Calculate similarity matrix
    n = len(ja_sentences)
    m = len(en_sentences)
    similarity_matrix = np.zeros((n, m))
    
    logger.info("Building similarity matrix...")
    for i in tqdm(range(n)):
        for j in range(m):
            # Calculate length-based similarity
            ja_len = len(ja_sentences[i])
            en_len = len(en_sentences[j])
            len_ratio = min(ja_len, en_len) / max(ja_len, en_len) if max(ja_len, en_len) > 0 else 0
            
            # For Japanese to English, we expect a certain character ratio
            # Typical ratio multiplier (can be adjusted)
            ja_en_ratio_multiplier = 0.4  # Japanese text is often shorter in character count
            expected_en_len = ja_len / ja_en_ratio_multiplier
            ratio_diff = abs(en_len - expected_en_len) / expected_en_len if expected_en_len > 0 else 1
            ratio_similarity = max(0, 1 - ratio_diff)
            
            # Simple character-based similarity for numbers and named entities
            ja_nums = set(re.findall(r'\d+', ja_sentences[i]))
            en_nums = set(re.findall(r'\d+', en_sentences[j]))
            num_overlap = len(ja_nums.intersection(en_nums)) / max(len(ja_nums.union(en_nums)), 1) if ja_nums or en_nums else 0
            
            # Special markers similarity (for headings and tables)
            special_similarity = 0
            if ('[HEADING' in ja_sentences[i] and '[HEADING' in en_sentences[j]):
                special_similarity = 1.0
            elif ('[TABLE]' in ja_sentences[i] and '[TABLE]' in en_sentences[j]):
                special_similarity = 1.0
            
            # Calculate final similarity score (weighted)
            similarity_matrix[i, j] = (0.5 * ratio_similarity + 
                                       0.3 * num_overlap + 
                                       0.2 * special_similarity)
    
    # Use dynamic programming to find the best alignment path
    # Create a matrix to store the maximum sum of similarities
    dp = np.zeros((n + 1, m + 1))
    # Create a matrix to store the alignment decisions
    decisions = np.zeros((n + 1, m + 1), dtype=int)
    
    # Fill the DP matrix
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            # Three possible previous states: 1:1, 1:2, 2:1 alignment
            candidates = [
                dp[i-1, j-1] + similarity_matrix[i-1, j-1],  # 1:1 alignment
                dp[i-1, j-2] + (similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/2 if j >= 2 else -float('inf'),  # 1:2 alignment
                dp[i-2, j-1] + (similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/2 if i >= 2 else -float('inf')   # 2:1 alignment
            ]
            
            # Select the best decision
            best_decision = np.argmax(candidates)
            dp[i, j] = candidates[best_decision]
            decisions[i, j] = best_decision
    
    # Backtrack to find the alignment
    aligned_pairs = []
    i, j = n, m
    
    while i > 0 and j > 0:
        decision = decisions[i, j]
        
        if decision == 0:  # 1:1 alignment
            if similarity_matrix[i-1, j-1] >= similarity_threshold:
                aligned_pairs.append((ja_sentences[i-1], en_sentences[j-1]))
            i -= 1
            j -= 1
        elif decision == 1:  # 1:2 alignment
            # Combine two English sentences
            if j >= 2 and (similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/2 >= similarity_threshold:
                combined_en = en_sentences[j-2] + " " + en_sentences[j-1]
                aligned_pairs.append((ja_sentences[i-1], combined_en))
            i -= 1
            j -= 2
        else:  # 2:1 alignment
            # Combine two Japanese sentences
            if i >= 2 and (similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/2 >= similarity_threshold:
                combined_ja = ja_sentences[i-2] + " " + ja_sentences[i-1]
                aligned_pairs.append((combined_ja, en_sentences[j-1]))
            i -= 2
            j -= 1
    
    # Reverse the list to get the correct order
    aligned_pairs.reverse()
    
    logger.info(f"Successfully aligned {len(aligned_pairs)} sentence pairs")
    return aligned_pairs

def calculate_confidence_scores(aligned_pairs):
    """Calculate confidence scores for aligned sentence pairs."""
    confidence_scores = []
    
    for ja, en in aligned_pairs:
        # Length ratio confidence
        ja_len = len(ja)
        en_len = len(en)
        
        # For Japanese to English, we expect a certain character ratio
        ja_en_ratio_multiplier = 0.4  # Japanese text is often shorter in character count
        expected_en_len = ja_len / ja_en_ratio_multiplier
        ratio_diff = abs(en_len - expected_en_len) / expected_en_len if expected_en_len > 0 else 1
        ratio_confidence = max(0, 1 - ratio_diff)
        
        # Number and special character overlap confidence
        ja_nums = set(re.findall(r'\d+', ja))
        en_nums = set(re.findall(r'\d+', en))
        num_overlap = len(ja_nums.intersection(en_nums)) / max(len(ja_nums.union(en_nums)), 1) if ja_nums or en_nums else 1
        
        # Special markers confidence
        special_marker_match = 1.0 if ('[HEADING' in ja and '[HEADING' in en) or ('[TABLE]' in ja and '[TABLE]' in en) else 0.5
        
        # Overall confidence (weighted average)
        confidence = (0.5 * ratio_confidence + 0.3 * num_overlap + 0.2 * special_marker_match)
        confidence_scores.append(confidence)
    
    return confidence_scores

def create_parallel_corpus(ja_pdf_path, en_pdf_path, output_path, quality_review=False):
    """Create a Japanese-English parallel corpus from PDF files."""
    # Load spaCy models
    spacy_models = load_spacy_models()
    
    # Extract text from PDFs
    ja_text = extract_text_from_pdf(ja_pdf_path)
    en_text = extract_text_from_pdf(en_pdf_path)
    
    if not ja_text or not en_text:
        logger.error("Failed to extract text from one or both PDFs.")
        return False
    
    # Clean text
    ja_text_clean = clean_text(ja_text, 'ja')
    en_text_clean = clean_text(en_text, 'en')
    
    # Detect document structure
    ja_structure = detect_document_structure(ja_text_clean, 'ja')
    en_structure = detect_document_structure(en_text_clean, 'en')
    
    # Process text by structure type
    ja_sentences = []
    en_sentences = []
    
    logger.info("Processing Japanese document structure...")
    for element in tqdm(ja_structure, desc="Processing Japanese structure"):
        if element['type'] == 'heading':
            # Add heading with a special marker
            ja_sentences.append(f"[HEADING:{element['level']}] {element['text']}")
        elif element['type'] == 'table':
            # Add table with a special marker
            ja_sentences.append(f"[TABLE] {element['text']}")
        else:  # paragraph or other text
            # Segment paragraphs into sentences
            paragraph_sents = segment_into_sentences(element['text'], 'ja', spacy_models)
            ja_sentences.extend(paragraph_sents)
    
    logger.info("Processing English document structure...")
    for element in tqdm(en_structure, desc="Processing English structure"):
        if element['type'] == 'heading':
            # Add heading with a special marker
            en_sentences.append(f"[HEADING:{element['level']}] {element['text']}")
        elif element['type'] == 'table':
            # Add table with a special marker
            en_sentences.append(f"[TABLE] {element['text']}")
        else:  # paragraph or other text
            # Segment paragraphs into sentences
            paragraph_sents = segment_into_sentences(element['text'], 'en', spacy_models)
            en_sentences.extend(paragraph_sents)
    
    # Align sentences
    aligned_pairs = align_sentences_dynamic(ja_sentences, en_sentences)
    
    # Create DataFrame
    df = pd.DataFrame(aligned_pairs, columns=['Japanese', 'English'])
    
    # Save raw alignment results
    raw_output = output_path.replace('.csv', '_raw.csv')
    df.to_csv(raw_output, index=False, encoding='utf-8')
    logger.info(f"Saved raw alignment results to {raw_output}")
    
    # Quality review (optional)
    if quality_review:
        logger.info("Performing quality review...")
        
        # Calculate confidence scores
        confidence_scores = calculate_confidence_scores(aligned_pairs)
        
        # Add confidence scores to the DataFrame
        df['Confidence'] = confidence_scores
        
        # Filter by confidence threshold
        high_confidence_pairs = df[df['Confidence'] >= 0.7]
        medium_confidence_pairs = df[(df['Confidence'] >= 0.4) & (df['Confidence'] < 0.7)]
        low_confidence_pairs = df[df['Confidence'] < 0.4]
        
        # Save filtered results
        high_conf_output = output_path.replace('.csv', '_high_conf.csv')
        med_conf_output = output_path.replace('.csv', '_med_conf.csv')
        low_conf_output = output_path.replace('.csv', '_low_conf.csv')
        
        high_confidence_pairs.to_csv(high_conf_output, index=False, encoding='utf-8')
        medium_confidence_pairs.to_csv(med_conf_output, index=False, encoding='utf-8')
        low_confidence_pairs.to_csv(low_conf_output, index=False, encoding='utf-8')
        
        logger.info(f"Saved high confidence pairs ({len(high_confidence_pairs)}) to {high_conf_output}")
        logger.info(f"Saved medium confidence pairs ({len(medium_confidence_pairs)}) to {med_conf_output}")
        logger.info(f"Saved low confidence pairs ({len(low_confidence_pairs)}) to {low_conf_output}")
        
        # Use high confidence pairs for the final output
        df = high_confidence_pairs
    
    # Clean up the special markers from the final output
    df['Japanese'] = df['Japanese'].str.replace(r'\[HEADING:\d+\]\s*', '', regex=True)
    df['Japanese'] = df['Japanese'].str.replace(r'\[TABLE\]\s*', '', regex=True)
    df['English'] = df['English'].str.replace(r'\[HEADING:\d+\]\s*', '', regex=True)
    df['English'] = df['English'].str.replace(r'\[TABLE\]\s*', '', regex=True)
    
    # Save to final CSV
    logger.info(f"Saving {len(df)} aligned sentence pairs to {output_path}")
    df.to_csv(output_path, index=False, encoding='utf-8')
    
    logger.info("Parallel corpus creation completed successfully!")
    return True

def main():
    parser = argparse.ArgumentParser(description='Create a Japanese-English parallel corpus from PDF documents.')
    parser.add_argument('--ja-pdf', required=True, help='Path to the Japanese PDF file')
    parser.add_argument('--en-pdf', required=True, help='Path to the English PDF file')
    parser.add_argument('--output', required=True, help='Path to the output CSV file')
    parser.add_argument('--quality-review', action='store_true', help='Perform quality review and confidence scoring')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Set logging level based on verbose flag
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    create_parallel_corpus(args.ja_pdf, args.en_pdf, args.output, args.quality_review)

if __name__ == "__main__":
    main()
