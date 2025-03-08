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
    """
    Extract text from a PDF file with advanced options for handling columns, tables, 
    and proper text flow. This function uses multiple extraction techniques to ensure
    the best possible text extraction for sentence segmentation.
    """
    logger.info(f"Extracting text from: {pdf_path}")
    all_text = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(tqdm(pdf.pages, desc="Reading pages")):
                page_text = ""
                
                # First attempt: Extract tables if requested
                if detect_tables:
                    try:
                        tables = page.extract_tables()
                        for table in tables:
                            if table and any(cell for row in table for cell in row):
                                # Convert table to text representation
                                table_text = "\n".join([" | ".join([str(cell) if cell else "" for cell in row]) for row in table])
                                page_text += "[TABLE_START]\n" + table_text + "\n[TABLE_END]\n\n"
                    except Exception as e:
                        logger.warning(f"Table extraction failed on page {page_num+1}: {e}")
                
                # Main text extraction
                main_text = ""
                
                # Second attempt: Try column detection if requested
                if detect_columns:
                    try:
                        # Try to detect columns using word positions
                        words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
                        
                        if words:
                            # Group words by line using y-position (top to bottom)
                            words_by_line = {}
                            for word in words:
                                # Round y-position to group closely positioned words
                                y1 = round(word['top'] / 5) * 5  # Group within 5 pixels
                                if y1 not in words_by_line:
                                    words_by_line[y1] = []
                                words_by_line[y1].append(word)
                            
                            # Define a function to detect columns in a line
                            def detect_line_columns(words):
                                if len(words) < 4:  # Too few words to have multiple columns
                                    return [words]
                                
                                # Calculate gaps between words
                                sorted_words = sorted(words, key=lambda w: w['x0'])
                                gaps = []
                                for i in range(len(sorted_words) - 1):
                                    curr_word_end = sorted_words[i]['x1']
                                    next_word_start = sorted_words[i+1]['x0']
                                    gap = next_word_start - curr_word_end
                                    gaps.append((i, gap))
                                
                                # Find significant gaps (potential column separators)
                                if gaps:
                                    avg_gap = sum(g[1] for g in gaps) / len(gaps)
                                    std_gap = (sum((g[1] - avg_gap) ** 2 for g in gaps) / len(gaps)) ** 0.5
                                    
                                    # Identify large gaps (column separators)
                                    threshold = max(avg_gap + 1.5 * std_gap, 20)  # Higher threshold for columns
                                    column_breaks = sorted([g[0] for g in gaps if g[1] > threshold])
                                    
                                    if column_breaks:
                                        # Split words into columns
                                        columns = []
                                        start_idx = 0
                                        for break_idx in column_breaks:
                                            columns.append(sorted_words[start_idx:break_idx+1])
                                            start_idx = break_idx + 1
                                        columns.append(sorted_words[start_idx:])
                                        return columns
                                
                                # No columns detected
                                return [sorted_words]
                            
                            # Process each line with column detection
                            for y, line_words in sorted(words_by_line.items()):
                                columns = detect_line_columns(line_words)
                                
                                for column_words in columns:
                                    # Sort words left to right within each column
                                    sorted_column_words = sorted(column_words, key=lambda w: w['x0'])
                                    column_text = " ".join(w['text'] for w in sorted_column_words)
                                    if column_text.strip():
                                        main_text += column_text + "\n"
                    except Exception as e:
                        logger.warning(f"Column detection failed on page {page_num+1}: {e}")
                
                # Third attempt: If column detection failed or wasn't requested, use standard extraction
                if not main_text.strip():
                    try:
                        extracted = page.extract_text(x_tolerance=3, y_tolerance=3)
                        if extracted:
                            main_text = extracted
                    except Exception as e:
                        logger.warning(f"Standard text extraction failed on page {page_num+1}: {e}")
                
                # Add the main text after any tables
                page_text += main_text
                
                # Add page marker for debugging (can be removed in post-processing)
                page_text += f"\n[PAGE_END_{page_num+1}]\n\n"
                
                all_text.append(page_text)
        
        # Combine all pages
        full_text = "\n".join(all_text).strip()
        
        # Post-processing
        # 1. Remove page markers (used only for debugging)
        full_text = re.sub(r'\[PAGE_END_\d+\]\s*', '\n\n', full_text)
        
        # 2. Ensure table markers are properly spaced
        full_text = re.sub(r'\[TABLE_START\]\s*', '[TABLE_START]\n', full_text)
        full_text = re.sub(r'\s*\[TABLE_END\]', '\n[TABLE_END]', full_text)
        
        # 3. Fix common PDF text flow issues
        # Fix broken words at line endings (but preserve paragraph breaks)
        full_text = re.sub(r'([a-zA-Z0-9])-\n([a-zA-Z0-9])', r'\1\2', full_text)
        
        # 4. Handle consecutive hyphens (common in some PDFs)
        full_text = re.sub(r'(\w)- -(\w)', r'\1-\2', full_text)
        
        logger.info(f"Extracted {len(full_text)} characters from {pdf_path}")
        return full_text
    
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        return None

def clean_text(text, language):
    """
    Clean the extracted text to improve sentence segmentation and alignment.
    Handles common PDF extraction issues and removes headers, footers, page numbers, etc.
    """
    if not text:
        return text
    
    logger.info(f"Cleaning {language} text...")
    
    # Convert all whitespace to single spaces first
    text = re.sub(r'\s+', ' ', text)
    
    # Preserve line breaks when they indicate paragraph breaks
    text = re.sub(r'\n\s*\n', '\n\n', text)  # Keep paragraph breaks
    
    # Remove common PDF artifacts
    patterns_to_remove = [
        r'Page \d+ of \d+',
        r'Copyright © \d{4}.*?\n',
        r'^\s*\d+\s*

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
,  # Standalone page numbers
        r'^\s*[ivxlcdm]+\s*

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
,  # Roman numerals (often page numbers)
        r'^\s*[A-Z]\s*

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
,  # Single letters (often section markers)
    ]
    
    for pattern in patterns_to_remove:
        text = re.sub(pattern, '', text, flags=re.MULTILINE)
    
    # Remove headers and footers (likely to appear at top/bottom of pages)
    # This is a heuristic approach - headers/footers often appear on multiple pages
    lines = text.split('\n')
    
    # Find repeated lines that might be headers/footers
    line_counts = {}
    for line in lines:
        line = line.strip()
        if line and len(line) < 100:  # Headers/footers are usually short
            line_counts[line] = line_counts.get(line, 0) + 1
    
    # Identify potential headers/footers (appear on multiple pages)
    page_count = text.count('\f') + 1  # Form feed character indicates page breaks
    potential_headers_footers = []
    
    for line, count in line_counts.items():
        # If a line appears on more than 20% of pages and isn't just a number or date
        if (count > max(3, page_count * 0.2) and 
            not re.match(r'^\d+

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
, line) and 
            not re.match(r'^\d{1,2}/\d{1,2}/\d{2,4}

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
, line)):
            potential_headers_footers.append(line)
    
    # Remove identified headers/footers
    for header_footer in potential_headers_footers:
        text = text.replace(header_footer, '')
    
    # Language-specific cleaning
    if language.lower() == 'ja':
        # Remove unnecessary spaces in Japanese text (Japanese doesn't use spaces between words)
        text = re.sub(r'([^\w\s])\s+', r'\1', text)
        text = re.sub(r'\s+([^\w\s])', r'\1', text)
        
        # Fix common Japanese PDF extraction issues
        # Connect broken text that should be continuous
        text = re.sub(r'([^。！？、」）】]) *\n *([^\n「（【])', r'\1\2', text)
    else:  # English
        # Fix hyphenation at line breaks
        text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', text)
        
        # Fix broken sentences (when no punctuation before line break)
        text = re.sub(r'([a-z,;:])\s*\n\s*([a-z])', r'\1 \2', text)
    
    # Final cleanup
    # Remove excess whitespace again after all replacements
    text = re.sub(r' +', ' ', text)
    text = re.sub(r'\n\s+', '\n', text)
    
    # Ensure text starts and ends cleanly
    text = text.strip()
    
    return text

def detect_document_structure(text, language):
    """
    Detect headings, sections, and other structural elements.
    If structure detection fails, split text into manageable chunks to ensure sentence segmentation.
    """
    if not text:
        return []
    
    logger.info(f"Detecting document structure in {language} text...")
    
    # Split text into lines
    lines = text.split('\n')
    structured_elements = []
    
    # Heading detection patterns
    heading_patterns = [
        (r'^#{1,6}\s+(.+)

def segment_into_sentences(text, language, spacy_models):
    """
    Split text into sentences using a robust combination of spaCy and rule-based methods.
    This function will always attempt to split text into sentences even if spaCy fails.
    """
    if not text or len(text.strip()) == 0:
        return []
    
    # Track method used for logging
    method_used = "unknown"
    
    # Normalize text first to handle common PDF text issues
    # This improves segmentation regardless of which method we use
    normalized_text = text
    
    # Fix common PDF extraction issues
    # 1. Normalize line breaks
    normalized_text = re.sub(r'\r\n', '\n', normalized_text)
    
    # 2. Remove hyphenation at line breaks
    if language.lower() == 'en':
        normalized_text = re.sub(r'(\w)-\n(\w)', r'\1\2', normalized_text)
        # Join lines without proper sentence endings
        normalized_text = re.sub(r'([^.!?])\n', r'\1 ', normalized_text)
    elif language.lower() == 'ja':
        # Japanese doesn't use hyphenation but may have broken lines
        normalized_text = re.sub(r'([^。！？])\n', r'\1', normalized_text)
    
    # First try using spaCy for sentence segmentation (most accurate method)
    spacy_sentences = []
    try:
        # Limit text size to avoid memory issues with spaCy
        # For very large texts, we'll use rule-based methods
        if len(normalized_text) < 100000:  # 100KB limit
            if language.lower() == 'ja':
                doc = spacy_models['ja'](normalized_text)
            else:
                doc = spacy_models['en'](normalized_text)
            
            # Extract sentences from spaCy
            spacy_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
            
            # If spaCy found sentences, process them further
            if spacy_sentences:
                method_used = "spaCy"
                logger.debug(f"spaCy found {len(spacy_sentences)} sentences")
        else:
            logger.warning(f"Text too large for spaCy ({len(normalized_text)} chars). Using rule-based segmentation.")
    except Exception as e:
        logger.warning(f"Error in sentence segmentation using spaCy: {str(e)}")
    
    # If spaCy worked, process the results further for better accuracy
    if spacy_sentences:
        # Additional processing for sentences detected by spaCy
        processed_sentences = []
        
        if language.lower() == 'ja':
            # Process Japanese sentences from spaCy
            for sent in spacy_sentences:
                # Some Japanese might still need splitting on end punctuation
                sub_parts = re.split(r'(。|！|？)', sent)
                
                i = 0
                temp_sent = ""
                while i < len(sub_parts):
                    if i + 1 < len(sub_parts) and sub_parts[i+1] in ['。', '！', '？']:
                        # Complete sentence with punctuation
                        temp_sent = sub_parts[i] + sub_parts[i+1]
                        if temp_sent.strip():
                            processed_sentences.append(temp_sent.strip())
                        temp_sent = ""
                        i += 2
                    else:
                        # Incomplete sentence or ending
                        temp_sent += sub_parts[i]
                        i += 1
                
                # Add any remaining content
                if temp_sent.strip():
                    processed_sentences.append(temp_sent.strip())
                    
        else:  # English
            # Process English sentences from spaCy
            for sent in spacy_sentences:
                if len(sent) > 200:  # Very long sentence, might need further splitting
                    # Try to split on clear sentence boundaries
                    potential_splits = re.split(r'([.!?])\s+(?=[A-Z])', sent)
                    
                    i = 0
                    temp_sent = ""
                    while i < len(potential_splits):
                        if i + 1 < len(potential_splits) and potential_splits[i+1] in ['.', '!', '?']:
                            # Complete sentence with punctuation
                            temp_sent = potential_splits[i] + potential_splits[i+1]
                            if temp_sent.strip():
                                processed_sentences.append(temp_sent.strip())
                            temp_sent = ""
                            i += 2
                        else:
                            # Part without ending punctuation
                            temp_sent += potential_splits[i]
                            i += 1
                    
                    # Add any remaining content
                    if temp_sent.strip():
                        processed_sentences.append(temp_sent.strip())
                else:
                    # Regular sentence, keep as is
                    processed_sentences.append(sent)
                
        # Use processed spaCy results if they exist, otherwise continue to rule-based
        if processed_sentences:
            return [s for s in processed_sentences if s.strip()]
    
    # If spaCy failed or found no sentences, use rule-based segmentation (always works)
    method_used = "rule-based"
    rule_based_sentences = []
    
    if language.lower() == 'ja':
        # Japanese rule-based sentence segmentation
        # Split on Japanese sentence endings (。!?)
        
        # First handle quotes and parentheses to avoid splitting inside them
        # Add markers to help with sentence splitting while preserving quotes
        quote_pattern = r'「(.+?)」'
        normalized_text = re.sub(quote_pattern, lambda m: '「' + m.group(1).replace('。', '<PERIOD_MARKER>') + '」', normalized_text)
        
        # Now split on sentence endings not inside quotes
        parts = re.split(r'([。！？])', normalized_text)
        
        # Recombine parts into complete sentences
        i = 0
        current_sentence = ""
        while i < len(parts):
            current_part = parts[i]
            
            # If this is text followed by punctuation
            if i + 1 < len(parts) and parts[i+1] in ['。', '！', '？']:
                current_sentence += current_part + parts[i+1]
                
                # Check if the sentence is complete (not inside quotes/parentheses)
                # This is a simple check - a complete solution would need a stack
                if (current_sentence.count('「') == current_sentence.count('」') and
                    current_sentence.count('(') == current_sentence.count(')') and
                    current_sentence.count('（') == current_sentence.count('）')):
                    # Restore original periods inside quotes
                    current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
                    rule_based_sentences.append(current_sentence.strip())
                    current_sentence = ""
                
                i += 2
            else:
                # Text without punctuation, just add to current sentence
                current_sentence += current_part
                i += 1
        
        # Add any remaining text as a sentence
        if current_sentence.strip():
            # Restore original periods inside quotes
            current_sentence = current_sentence.replace('<PERIOD_MARKER>', '。')
            rule_based_sentences.append(current_sentence.strip())
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced Japanese sentence splitting failed, falling back to basic splitting")
            # Simple split on sentence endings
            simple_parts = re.split(r'([。！？]+)', normalized_text)
            simple_sentences = []
            
            i = 0
            while i < len(simple_parts) - 1:
                if i + 1 < len(simple_parts) and re.match(r'[。！？]+', simple_parts[i+1]):
                    simple_sentences.append((simple_parts[i] + simple_parts[i+1]).strip())
                    i += 2
                else:
                    if simple_parts[i].strip():
                        simple_sentences.append(simple_parts[i].strip())
                    i += 1
            
            # Add the last part if it exists
            if i < len(simple_parts) and simple_parts[i].strip():
                simple_sentences.append(simple_parts[i].strip())
                
            rule_based_sentences = simple_sentences
            
    else:  # English
        # English rule-based sentence segmentation
        
        # Handle common abbreviations that contain periods
        common_abbr = r'\b(Mr|Mrs|Ms|Dr|Prof|St|Jr|Sr|Inc|Ltd|Co|Corp|vs|etc|Jan|Feb|Mar|Apr|Aug|Sept|Oct|Nov|Dec|U\.S|a\.m|p\.m)\.'
        # Temporarily replace periods in abbreviations
        normalized_text = re.sub(common_abbr, r'\1<PERIOD_MARKER>', normalized_text)
        
        # Split on sentence boundaries: period, exclamation, question mark followed by space and capital
        splits = re.split(r'([.!?])\s+(?=[A-Z])', normalized_text)
        
        # Recombine into complete sentences
        i = 0
        current_sentence = ""
        while i < len(splits):
            if i + 1 < len(splits) and splits[i+1] in ['.', '!', '?']:
                # Complete sentence with punctuation
                current_sentence = splits[i] + splits[i+1]
                # Restore original periods in abbreviations
                current_sentence = current_sentence.replace('<PERIOD_MARKER>', '.')
                if current_sentence.strip():
                    rule_based_sentences.append(current_sentence.strip())
                i += 2
            else:
                # Text without matched punctuation
                current_part = splits[i]
                # Restore original periods in abbreviations
                current_part = current_part.replace('<PERIOD_MARKER>', '.')
                if current_part.strip():
                    rule_based_sentences.append(current_part.strip())
                i += 1
        
        # If we still don't have sentences, do a simpler split
        if not rule_based_sentences:
            logger.warning("Advanced English sentence splitting failed, falling back to basic splitting")
            # Split on any sequence of punctuation followed by whitespace
            simple_sentences = re.split(r'(?<=[.!?])\s+', normalized_text)
            rule_based_sentences = [s.strip() for s in simple_sentences if s.strip()]
    
    # Final filtering and sentence length check
    final_sentences = []
    for sentence in rule_based_sentences:
        # Skip very short "sentences" that are likely not real sentences
        min_chars = 2 if language.lower() == 'ja' else 4
        if len(sentence) >= min_chars:
            final_sentences.append(sentence)
        
    logger.debug(f"Rule-based method found {len(final_sentences)} sentences")
    
    # Last resort: if we still have no sentences, just return the text as one sentence
    if not final_sentences and normalized_text.strip():
        logger.warning("All sentence segmentation methods failed. Treating text as a single sentence.")
        return [normalized_text.strip()]
    
    return final_sentences

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
, 'markdown'),  # Markdown headings
        (r'^(\d+\.)+\s+(.+)

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
, 'numbered'),  # Numbered headings like "1.2.3 Title"
        (r'^(Chapter|Section|Part|章|節)\s*\d*\s*:?\s*(.+)

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
, 'labeled'),  # Labeled headings in English and Japanese
    ]
    
    current_paragraph = []
    in_table = False
    table_content = []
    
    # Process line by line
    for line in lines:
        line = line.strip()
        if not line:
            # Process completed paragraph
            if current_paragraph:
                paragraph_text = ' '.join(current_paragraph)
                structured_elements.append({
                    'text': paragraph_text,
                    'type': 'paragraph',
                    'level': 0
                })
                current_paragraph = []
            
            # Process completed table
            if in_table and table_content:
                structured_elements.append({
                    'text': '\n'.join(table_content),
                    'type': 'table',
                    'level': 0
                })
                table_content = []
                in_table = False
            continue
        
        # Check if line is a table row
        if '|' in line and line.count('|') >= 2:
            if not in_table:
                # Process any pending paragraph before starting table
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
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
        for pattern, h_type in heading_patterns:
            match = re.match(pattern, line)
            if match:
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                is_heading = True
                if h_type == 'markdown':
                    heading_level = line.index(' ')
                    heading_text = match.group(1)
                elif h_type == 'numbered':
                    heading_level = line.count('.')
                    heading_text = match.group(2)
                else:  # labeled
                    heading_level = 1
                    heading_text = match.group(2)
                
                structured_elements.append({
                    'text': heading_text,
                    'type': 'heading',
                    'level': heading_level
                })
                break
        
        if not is_heading:
            # Check if line appears to be a title or heading based on length and capitalization
            if (len(line) < 100 and (line.isupper() or line.istitle()) or 
                (language == 'ja' and re.match(r'^[\u3000-\u303F\u4E00-\u9FAF\u3040-\u309F\u30A0-\u30FF]{1,20}

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
, line))):
                
                # Process any pending paragraph before heading
                if current_paragraph:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
                
                structured_elements.append({
                    'text': line,
                    'type': 'heading',
                    'level': 1
                })
            else:
                # Add to current paragraph
                current_paragraph.append(line)
                
                # If paragraph gets too long, process it as a chunk to avoid memory issues
                # and ensure we don't end up with just one giant paragraph
                if len(' '.join(current_paragraph)) > 2000:
                    paragraph_text = ' '.join(current_paragraph)
                    structured_elements.append({
                        'text': paragraph_text,
                        'type': 'paragraph',
                        'level': 0
                    })
                    current_paragraph = []
    
    # Process any remaining content
    if current_paragraph:
        paragraph_text = ' '.join(current_paragraph)
        structured_elements.append({
            'text': paragraph_text,
            'type': 'paragraph',
            'level': 0
        })
    
    if in_table and table_content:
        structured_elements.append({
            'text': '\n'.join(table_content),
            'type': 'table',
            'level': 0
        })
    
    # If structure detection failed (only one huge element), force chunking
    if len(structured_elements) <= 1 and text:
        logger.warning(f"Structure detection found only {len(structured_elements)} elements. Forcing chunking...")
        
        # Clear the elements
        structured_elements = []
        
        # Split the text into chunks of approximately 1000 characters each
        chunk_size = 1000
        # Split by newlines first to preserve some structure
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if not paragraph.strip():
                continue
                
            # If paragraph is very long, chunk it further
            if len(paragraph) > chunk_size:
                # For Japanese, try to split on sentence endings
                if language == 'ja':
                    # Split on Japanese sentence endings
                    chunks = re.split(r'([。！？]+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
                else:  # English
                    # Split on English sentence endings
                    chunks = re.split(r'([.!?]+\s+)', paragraph)
                    
                    current_chunk = []
                    current_length = 0
                    
                    # Rebuild sentences ensuring punctuation stays with sentence
                    for i in range(0, len(chunks), 2):
                        sentence = chunks[i]
                        # Add punctuation if available
                        if i + 1 < len(chunks):
                            sentence += chunks[i + 1]
                        
                        if current_length + len(sentence) > chunk_size and current_chunk:
                            # Save current chunk
                            structured_elements.append({
                                'text': ''.join(current_chunk),
                                'type': 'paragraph',
                                'level': 0
                            })
                            current_chunk = [sentence]
                            current_length = len(sentence)
                        else:
                            current_chunk.append(sentence)
                            current_length += len(sentence)
                    
                    # Add any remaining content
                    if current_chunk:
                        structured_elements.append({
                            'text': ''.join(current_chunk),
                            'type': 'paragraph',
                            'level': 0
                        })
            else:
                # Short paragraph, add as is
                structured_elements.append({
                    'text': paragraph,
                    'type': 'paragraph',
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
