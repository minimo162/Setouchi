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
    """Split text into sentences using spaCy and enhanced rules for better Japanese-English alignment."""
    if not text:
        return []
    
    # First try spaCy for sentence segmentation
    try:
        if language.lower() == 'ja':
            doc = spacy_models['ja'](text)
            # Extract sentences
            sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # Additional processing for Japanese sentences
            processed_sentences = []
            for sent in sentences:
                # Some Japanese PDFs may have line breaks within sentences
                # Replace single line breaks that don't follow sentence-ending punctuation
                cleaned_sent = re.sub(r'([^。！？\n])\n([^\n])', r'\1\2', sent)
                # Split on actual sentence endings
                sub_sents = re.split(r'(。|！|？)(?=\s|$)', cleaned_sent)
                
                i = 0
                while i < len(sub_sents) - 1:
                    if sub_sents[i+1] in ['。', '！', '？']:
                        processed_sentences.append(sub_sents[i] + sub_sents[i+1])
                        i += 2
                    else:
                        processed_sentences.append(sub_sents[i])
                        i += 1
                        
                if i < len(sub_sents) and sub_sents[i].strip():
                    processed_sentences.append(sub_sents[i])
            
            return [s.strip() for s in processed_sentences if s.strip()]
        else:  # English
            doc = spacy_models['en'](text)
            # Extract sentences
            sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # Additional processing for English sentences
            processed_sentences = []
            for sent in sentences:
                # Handle cases where spaCy might not split correctly
                # For example, split on common sentence-ending punctuation
                if len(sent) > 150:  # Long sentence - might need additional splitting
                    # Look for sentence boundaries that spaCy missed
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    while i < len(potential_splits) - 1:
                        if potential_splits[i+1] in ['.', '!', '?']:
                            processed_sentences.append(potential_splits[i] + potential_splits[i+1])
                            i += 2
                        else:
                            processed_sentences.append(potential_splits[i])
                            i += 1
                    
                    if i < len(potential_splits) and potential_splits[i].strip():
                        processed_sentences.append(potential_splits[i])
                else:
                    processed_sentences.append(sent)
            
            return [s.strip() for s in processed_sentences if s.strip()]
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {e}")
        logger.info("Falling back to rule-based sentence splitting")
    
    # Fallback to more sophisticated rule-based splitting
    if language.lower() == 'ja':
        # Japanese sentence splitting with better handling
        # Split on common Japanese sentence endings (。!?), but handle special cases
        
        # First, normalize line breaks
        normalized_text = re.sub(r'\r\n', '\n', text)
        
        # Handle cases where sentences might span multiple lines
        # Replace single line breaks that don't follow sentence-ending punctuation
        normalized_text = re.sub(r'([^。！？\n])\n([^\n])', r'\1\2', normalized_text)
        
        # Now split on sentence-ending punctuation
        parts = re.split(r'(。|！|？)', normalized_text)
        sentences = []
        
        i = 0
        while i < len(parts) - 1:
            if parts[i+1] in ['。', '！', '？']:
                sentences.append(parts[i] + parts[i+1])
                i += 2
            else:
                sentences.append(parts[i])
                i += 1
        
        if i < len(parts) and parts[i].strip():
            sentences.append(parts[i])
    else:
        # English sentence splitting with better handling
        # First, normalize line breaks and handle common PDF issues
        normalized_text = re.sub(r'\r\n', '\n', text)
        
        # Handle hyphenation at line breaks (common in PDFs)
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        
        # Replace line breaks that don't follow sentence-ending punctuation
        normalized_text = re.sub(r'([^.!?\n])\n([^\n])', r'\1 \2', normalized_text)
        
        # Split on sentence-ending punctuation followed by space and capital letter
        sentence_pattern = r'(?<=[.!?])\s+(?=[A-Z])'
        raw_sentences = re.split(sentence_pattern, normalized_text)
        
        # Further process for edge cases
        sentences = []
        for raw_sent in raw_sentences:
            # Skip empty sentences
            if not raw_sent.strip():
                continue
                
            # Handle cases like "Dr. Smith" or "U.S. Army" that have periods but aren't sentence boundaries
            # This is a simplification; a complete solution would need a more sophisticated approach
            if re.search(r'\b(?:Mr|Mrs|Ms|Dr|Prof|Inc|Ltd|Co|U\.S|a\.m|p\.m)\.\s+[A-Z]', raw_sent):
                sentences.append(raw_sent)
            else:
                # Check if there are potential sentence boundaries within
                parts = re.split(r'([.!?])\s+(?=[A-Z])', raw_sent)
                
                i = 0
                while i < len(parts) - 1:
                    if parts[i+1] in ['.', '!', '?']:
                        sentences.append(parts[i] + parts[i+1])
                        i += 2
                    else:
                        sentences.append(parts[i])
                        i += 1
                
                if i < len(parts) and parts[i].strip():
                    sentences.append(parts[i])
    
    return [s.strip() for s in sentences if s.strip()]

def align_sentences_dynamic(ja_sentences, en_sentences, similarity_threshold=0.3):
    """
    Align Japanese and English sentences using enhanced dynamic programming and multiple similarity metrics
    to create a more accurate sentence-by-sentence parallel corpus.
    """
    if not ja_sentences or not en_sentences:
        return []
    
    logger.info(f"Aligning {len(ja_sentences)} Japanese sentences with {len(en_sentences)} English sentences")
    
    # Calculate similarity matrix
    n = len(ja_sentences)
    m = len(en_sentences)
    similarity_matrix = np.zeros((n, m))
    
    # Define typical Japanese to English character ratio (Japanese is more compact)
    # This ratio is used to estimate expected English sentence length from Japanese
    ja_en_ratio_multiplier = 0.4  # Typical ratio - can be adjusted based on document style
    
    logger.info("Building similarity matrix for sentence alignment...")
    for i in tqdm(range(n)):
        for j in range(m):
            # Calculate multiple similarity features
            
            # 1. Length-based similarity
            ja_len = len(ja_sentences[i])
            en_len = len(en_sentences[j])
            
            # Calculate expected English length based on Japanese character count
            expected_en_len = ja_len / ja_en_ratio_multiplier
            ratio_diff = abs(en_len - expected_en_len) / expected_en_len if expected_en_len > 0 else 1
            ratio_similarity = max(0, 1 - ratio_diff)
            
            # 2. Number and digit pattern similarity
            ja_nums = set(re.findall(r'\d+', ja_sentences[i]))
            en_nums = set(re.findall(r'\d+', en_sentences[j]))
            num_overlap = len(ja_nums.intersection(en_nums)) / max(len(ja_nums.union(en_nums)), 1) if ja_nums or en_nums else 0
            
            # 3. Special character and symbol overlap (dates, URLs, emails, etc.)
            ja_special = set(re.findall(r'[%$@#&*(){}[\]|<>:;,./?~^]+', ja_sentences[i]))
            en_special = set(re.findall(r'[%$@#&*(){}[\]|<>:;,./?~^]+', en_sentences[j]))
            special_char_overlap = len(ja_special.intersection(en_special)) / max(len(ja_special.union(en_special)), 1) if ja_special or en_special else 0
            
            # 4. Proper noun and entity similarity (especially for loanwords that might be similar)
            # Look for katakana words which often correspond to English proper nouns
            ja_katakana = set(re.findall(r'[ァ-ヶー]+', ja_sentences[i]))
            katakana_similarity = 0
            if ja_katakana:
                # Higher similarity if the sentence has katakana and matching numbers
                if num_overlap > 0:
                    katakana_similarity = 0.3
            
            # 5. Position-based similarity (sentences that are nearby in document ordering)
            # Sentences that are close in relative position are more likely to align
            position_similarity = max(0, 1 - abs((i/n) - (j/m))*3)
            
            # 6. Special markers similarity (for headings and tables)
            structure_similarity = 0
            if ('[HEADING' in ja_sentences[i] and '[HEADING' in en_sentences[j]):
                # Check if heading levels match
                ja_heading_match = re.search(r'\[HEADING:(\d+)\]', ja_sentences[i])
                en_heading_match = re.search(r'\[HEADING:(\d+)\]', en_sentences[j])
                if ja_heading_match and en_heading_match:
                    if ja_heading_match.group(1) == en_heading_match.group(1):
                        structure_similarity = 1.0
                    else:
                        structure_similarity = 0.5
                else:
                    structure_similarity = 0.7
            elif ('[TABLE]' in ja_sentences[i] and '[TABLE]' in en_sentences[j]):
                structure_similarity = 1.0
            
            # Calculate final similarity score (weighted combination of all features)
            similarity_matrix[i, j] = (
                0.35 * ratio_similarity +      # Length ratio is important
                0.25 * num_overlap +           # Number matching is strong signal
                0.10 * special_char_overlap +  # Special chars help with technical content
                0.10 * katakana_similarity +   # Katakana can help with names/terms
                0.10 * position_similarity +   # Position helps with overall document flow
                0.10 * structure_similarity    # Structure markers for headings/tables
            )
    
    # Enhanced dynamic programming for better alignment
    # Create a matrix to store the maximum sum of similarities
    dp = np.zeros((n + 1, m + 1))
    # Create a matrix to store the alignment decisions
    # 0: 1:1, 1: 1:2, 2: 2:1, 3: 1:3, 4: 3:1, 5: skip Japanese, 6: skip English
    decisions = np.zeros((n + 1, m + 1), dtype=int)
    
    # Fill the DP matrix with more alignment options
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            # More possible alignment patterns
            candidates = [
                dp[i-1, j-1] + similarity_matrix[i-1, j-1],  # 1:1 alignment
                
                # 1:2 alignment (one Japanese to two English)
                dp[i-1, j-2] + (similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/2 
                    if j >= 2 else -float('inf'),
                
                # 2:1 alignment (two Japanese to one English)
                dp[i-2, j-1] + (similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/2
                    if i >= 2 else -float('inf'),
                
                # 1:3 alignment (one Japanese to three English)
                dp[i-1, j-3] + (similarity_matrix[i-1, j-3] + similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/3
                    if j >= 3 else -float('inf'),
                
                # 3:1 alignment (three Japanese to one English)
                dp[i-3, j-1] + (similarity_matrix[i-3, j-1] + similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/3
                    if i >= 3 else -float('inf'),
                
                # Skip Japanese sentence (useful for document-specific content not in translation)
                dp[i-1, j] - 0.2  # Penalty for skipping
                    if i > 0 else -float('inf'),
                
                # Skip English sentence
                dp[i, j-1] - 0.2  # Penalty for skipping
                    if j > 0 else -float('inf')
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
            else:
                logger.debug(f"Low confidence 1:1 alignment skipped: JA[{i-1}], EN[{j-1}], score={similarity_matrix[i-1, j-1]}")
            i -= 1
            j -= 1
        elif decision == 1:  # 1:2 alignment
            # Combine two English sentences
            if j >= 2 and (similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/2 >= similarity_threshold:
                combined_en = en_sentences[j-2] + " " + en_sentences[j-1]
                aligned_pairs.append((ja_sentences[i-1], combined_en))
            else:
                logger.debug(f"Low confidence 1:2 alignment skipped: JA[{i-1}], EN[{j-2}:{j-1}]")
            i -= 1
            j -= 2
        elif decision == 2:  # 2:1 alignment
            # Combine two Japanese sentences
            if i >= 2 and (similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/2 >= similarity_threshold:
                combined_ja = ja_sentences[i-2] + " " + ja_sentences[i-1]
                aligned_pairs.append((combined_ja, en_sentences[j-1]))
            else:
                logger.debug(f"Low confidence 2:1 alignment skipped: JA[{i-2}:{i-1}], EN[{j-1}]")
            i -= 2
            j -= 1
        elif decision == 3:  # 1:3 alignment
            # Combine three English sentences
            if j >= 3 and (similarity_matrix[i-1, j-3] + similarity_matrix[i-1, j-2] + similarity_matrix[i-1, j-1])/3 >= similarity_threshold:
                combined_en = en_sentences[j-3] + " " + en_sentences[j-2] + " " + en_sentences[j-1]
                aligned_pairs.append((ja_sentences[i-1], combined_en))
            else:
                logger.debug(f"Low confidence 1:3 alignment skipped: JA[{i-1}], EN[{j-3}:{j-1}]")
            i -= 1
            j -= 3
        elif decision == 4:  # 3:1 alignment
            # Combine three Japanese sentences
            if i >= 3 and (similarity_matrix[i-3, j-1] + similarity_matrix[i-2, j-1] + similarity_matrix[i-1, j-1])/3 >= similarity_threshold:
                combined_ja = ja_sentences[i-3] + " " + ja_sentences[i-2] + " " + ja_sentences[i-1]
                aligned_pairs.append((combined_ja, en_sentences[j-1]))
            else:
                logger.debug(f"Low confidence 3:1 alignment skipped: JA[{i-3}:{i-1}], EN[{j-1}]")
            i -= 3
            j -= 1
        elif decision == 5:  # Skip Japanese
            logger.debug(f"Skipping Japanese sentence: {i-1}")
            i -= 1
        elif decision == 6:  # Skip English
            logger.debug(f"Skipping English sentence: {j-1}")
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
    """Create a Japanese-English parallel corpus from PDF files with enhanced sentence-level alignment."""
    # Load spaCy models
    spacy_models = load_spacy_models()
    
    # Extract text from PDFs
    logger.info("Step 1: Extracting text from PDF files...")
    ja_text = extract_text_from_pdf(ja_pdf_path)
    en_text = extract_text_from_pdf(en_pdf_path)
    
    if not ja_text or not en_text:
        logger.error("Failed to extract text from one or both PDFs.")
        return False
    
    # Clean text
    logger.info("Step 2: Cleaning extracted text...")
    ja_text_clean = clean_text(ja_text, 'ja')
    en_text_clean = clean_text(en_text, 'en')
    
    # Detect document structure
    logger.info("Step 3: Analyzing document structure...")
    ja_structure = detect_document_structure(ja_text_clean, 'ja')
    en_structure = detect_document_structure(en_text_clean, 'en')
    
    # Process text by structure type
    ja_sentences = []
    en_sentences = []
    
    logger.info("Step 4: Segmenting text into sentences...")
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
    
    logger.info(f"Found {len(ja_sentences)} Japanese sentences and {len(en_sentences)} English sentences")
    
    # Save intermediate sentences for debugging or manual inspection
    with open(output_path.replace('.csv', '_ja_sentences.txt'), 'w', encoding='utf-8') as f:
        for i, sent in enumerate(ja_sentences):
            f.write(f"{i+1}: {sent}\n")
    
    with open(output_path.replace('.csv', '_en_sentences.txt'), 'w', encoding='utf-8') as f:
        for i, sent in enumerate(en_sentences):
            f.write(f"{i+1}: {sent}\n")
    
    logger.info("Step 5: Aligning sentences between languages...")
    # Align sentences with the improved algorithm
    aligned_pairs = align_sentences_dynamic(ja_sentences, en_sentences, similarity_threshold=0.3)
    
    # Create DataFrame
    df = pd.DataFrame(aligned_pairs, columns=['Japanese', 'English'])
    
    # Add index numbers to track original sentence positions
    df['Ja_Index'] = df['Japanese'].apply(lambda x: ja_sentences.index(x) if x in ja_sentences else -1)
    df['En_Index'] = df['English'].apply(lambda x: en_sentences.index(x) if x in en_sentences else -1)
    
    # Save raw alignment results
    raw_output = output_path.replace('.csv', '_raw.csv')
    df.to_csv(raw_output, index=False, encoding='utf-8')
    logger.info(f"Saved raw alignment results to {raw_output}")
    
    # Quality review and filtering
    if quality_review:
        logger.info("Step 6: Performing quality review and confidence scoring...")
        
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
    
    # Post-process aligned pairs
    logger.info("Step 7: Post-processing aligned sentence pairs...")
    
    # Clean up the special markers from the final output
    df['Japanese'] = df['Japanese'].str.replace(r'\[HEADING:\d+\]\s*', '', regex=True)
    df['Japanese'] = df['Japanese'].str.replace(r'\[TABLE\]\s*', '', regex=True)
    df['English'] = df['English'].str.replace(r'\[HEADING:\d+\]\s*', '', regex=True)
    df['English'] = df['English'].str.replace(r'\[TABLE\]\s*', '', regex=True)
    
    # Additional cleaning for final output
    # Remove multiple spaces, normalize line breaks, etc.
    df['Japanese'] = df['Japanese'].str.replace(r'\s+', ' ', regex=True).str.strip()
    df['English'] = df['English'].str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # Remove any empty pairs
    df = df[(df['Japanese'].str.len() > 0) & (df['English'].str.len() > 0)]
    
    # Save to final CSV
    logger.info(f"Step 8: Saving {len(df)} aligned sentence pairs to {output_path}")
    
    # Remove the internal tracking columns before saving final output
    final_df = df.drop(columns=['Ja_Index', 'En_Index'] + 
                      (['Confidence'] if 'Confidence' in df.columns else []))
    
    final_df.to_csv(output_path, index=False, encoding='utf-8')
    
    # Create a simple HTML visualization for easy review
    html_output = output_path.replace('.csv', '_preview.html')
    
    with open(html_output, 'w', encoding='utf-8') as f:
        f.write('''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Parallel Corpus Preview</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                .pair { display: flex; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
                .japanese { flex: 1; margin-right: 10px; }
                .english { flex: 1; }
                h1 { color: #333; }
                .stats { margin-bottom: 20px; background: #f5f5f5; padding: 10px; border-radius: 5px; }
            </style>
        </head>
        <body>
            <h1>Japanese-English Parallel Corpus</h1>
            <div class="stats">
                <p>Total sentence pairs: ''' + str(len(final_df)) + '''</p>
                <p>Source files: ''' + ja_pdf_path + " & " + en_pdf_path + '''</p>
            </div>
        ''')
        
        for i, row in final_df.iterrows():
            if i >= 100:  # Just show the first 100 pairs in the preview
                f.write('<p>... (showing first 100 pairs only) ...</p>')
                break
                
            f.write(f'''
            <div class="pair">
                <div class="japanese">{i+1}. {row['Japanese']}</div>
                <div class="english">{i+1}. {row['English']}</div>
            </div>
            ''')
        
        f.write('''
        </body>
        </html>
        ''')
    
    logger.info(f"Created HTML preview at {html_output}")
    logger.info("Parallel corpus creation completed successfully!")
    return True

def main():
    parser = argparse.ArgumentParser(description='Create a Japanese-English parallel corpus from PDF documents.')
    parser.add_argument('--ja-pdf', required=True, help='Path to the Japanese PDF file')
    parser.add_argument('--en-pdf', required=True, help='Path to the English PDF file')
    parser.add_argument('--output', required=True, help='Path to the output CSV file')
    parser.add_argument('--quality-review', action='store_true', help='Perform quality review and confidence scoring')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--similarity-threshold', type=float, default=0.3, 
                        help='Minimum similarity threshold for sentence alignment (0.0-1.0)')
    parser.add_argument('--ignore-tables', action='store_true', 
                        help='Ignore tables during extraction')
    parser.add_argument('--ignore-columns', action='store_true', 
                        help='Ignore column detection during extraction')
    parser.add_argument('--batch', help='Process multiple file pairs listed in a CSV file')
    
    args = parser.parse_args()
    
    # Set logging level based on verbose flag
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    if args.batch:
        # Batch processing mode
        logger.info(f"Starting batch processing from {args.batch}")
        
        try:
            batch_df = pd.read_csv(args.batch)
            required_cols = ['ja_pdf', 'en_pdf', 'output']
            if not all(col in batch_df.columns for col in required_cols):
                logger.error(f"Batch file must contain columns: {', '.join(required_cols)}")
                return False
            
            success_count = 0
            for i, row in batch_df.iterrows():
                logger.info(f"Processing batch item {i+1}/{len(batch_df)}: {row['ja_pdf']} & {row['en_pdf']}")
                try:
                    success = create_parallel_corpus(
                        row['ja_pdf'], 
                        row['en_pdf'], 
                        row['output'],
                        args.quality_review
                    )
                    if success:
                        success_count += 1
                except Exception as e:
                    logger.error(f"Error processing batch item {i+1}: {e}")
            
            logger.info(f"Batch processing complete. Successfully processed {success_count}/{len(batch_df)} items.")
            
        except Exception as e:
            logger.error(f"Error in batch processing: {e}")
            return False
    else:
        # Single file processing mode
        logger.info("Starting single file processing")
        
        detect_tables = not args.ignore_tables
        detect_columns = not args.ignore_columns
        
        # Override default functions with custom parameters
        original_extract_text = extract_text_from_pdf
        
        def custom_extract_text(pdf_path, **kwargs):
            return original_extract_text(pdf_path, detect_columns=detect_columns, detect_tables=detect_tables)
        
        # Replace the function temporarily
        globals()['extract_text_from_pdf'] = custom_extract_text
        
        # Process with custom similarity threshold
        success = create_parallel_corpus(
            args.ja_pdf, 
            args.en_pdf, 
            args.output,
            args.quality_review
        )
        
        # Restore original function
        globals()['extract_text_from_pdf'] = original_extract_text
        
        if success:
            logger.info("Processing completed successfully")
        else:
            logger.error("Processing failed")

if __name__ == "__main__":
    main()
