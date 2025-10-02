from pathlib import Path

path = Path('excel_copilot/tools/excel_tools.py')
text = path.read_text(encoding='utf-8')
old = "                existing_output_value = output_matrix[local_row][translation_col_index]\n                if translation_value != existing_output_value:\n                    output_matrix[local_row][translation_col_index] = translation_value\n                    output_dirty = True\n                if not writing_to_source_directly and overwrite_source:\n                    existing_source_value = source_matrix[local_row][col_idx]\n                    if translation_value != existing_source_value:\n                        source_matrix[local_row][col_idx] = translation_value\n                        source_dirty = True\n\n                any_translation = True\n\n                explanation_text = explanation_jp.strip()\n"
new = "                existing_output_value = output_matrix[local_row][translation_col_index]\n                if translation_value != existing_output_value:\n                    output_matrix[local_row][translation_col_index] = translation_value\n                    output_dirty = True\n                if not writing_to_source_directly and overwrite_source:\n                    existing_source_value = source_matrix[local_row][col_idx]\n                    if translation_value != existing_source_value:\n                        source_matrix[local_row][col_idx] = translation_value\n                        source_dirty = True\n\n                explanation_text = explanation_jp.strip()\n"
if old not in text:
    raise RuntimeError('Expected block not found when removing any_translation assignment')
text = text.replace(old, new, 1)

insertion_point = "                    quote_candidates = aggregated_quotes\n\n                    existing_output_value = output_matrix[local_row][translation_col_index]\n"
insertion_replacement = "                    quote_candidates = aggregated_quotes\n\n                    existing_output_value = output_matrix[local_row][translation_col_index]\n                    if translation_value != existing_output_value:\n                        output_matrix[local_row][translation_col_index] = translation_value\n                        output_dirty = True\n                    if not writing_to_source_directly and overwrite_source:\n                        existing_source_value = source_matrix[local_row][col_idx]\n                        if translation_value != existing_source_value:\n                            source_matrix[local_row][col_idx] = translation_value\n                            source_dirty = True\n
"
# First, remove duplicated lines inserted earlier (since we removed earlier ones). We'll adjust aggregator block to avoid duplication.
if insertion_point not in text:
    raise RuntimeError('Aggregator block not found for reinsertion')
text = text.replace(insertion_point, "                    quote_candidates = aggregated_quotes\n\n                    existing_output_value = output_matrix[local_row][translation_col_index]\n                    if translation_value != existing_output_value:\n                        output_matrix[local_row][translation_col_index] = translation_value\n                        output_dirty = True\n                    if not writing_to_source_directly and overwrite_source:\n                        existing_source_value = source_matrix[local_row][col_idx]\n                        if translation_value != existing_source_value:\n                            source_matrix[local_row][col_idx] = translation_value\n                            source_dirty = True\n\n", 1)

insert_after = "                    quote_candidates = aggregated_quotes\n\n                    existing_output_value = output_matrix[local_row][translation_col_index]\n                    if translation_value != existing_output_value:\n                        output_matrix[local_row][translation_col_index] = translation_value\n                        output_dirty = True\n                    if not writing_to_source_directly and overwrite_source:\n                        existing_source_value = source_matrix[local_row][col_idx]\n                        if translation_value != existing_source_value:\n                            source_matrix[local_row][col_idx] = translation_value\n                            source_dirty = True\n\n"
if insert_after not in text:
    raise RuntimeError('Failed to find location to insert any_translation assignment')
text = text.replace(insert_after, insert_after + "                any_translation = True\n\n", 1)

path.write_text(text, encoding='utf-8')
