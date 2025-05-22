import os
from nltk.corpus import wordnet
import re


def find_acronyms(word_app, base_definition_path, user_definition_path, context_range=5):
    # Determine base path for custom lists.
    # Uses the directory of the main definition_path for custom user lists.
    # The main definition_path is cached at:
    # os.path.join(os.path.expanduser("~"), ".doc_companion", "acronym_list.txt")
    # So, user_exclude.txt and user_include.txt should be placed in ~/.doc_companion/
    user_config_dir = os.path.dirname(user_definition_path)
    base_defs_map = _load_definitions_from_file(base_definition_path)
    user_defs_map = _load_definitions_from_file(user_definition_path)
    combined_defs_map = {**base_defs_map, **user_defs_map}
    defined_acronyms = set(combined_defs_map.keys()) # This set is used for 'word in defined_acronyms' checks
    exclude_filepath = os.path.join(user_config_dir, "user_exclude.txt")
    include_filepath = os.path.join(user_config_dir, "user_include.txt")

    # Default exclusion list
    DEFAULT_EXCLUDE = {
        'ON', 'BC', 'MB', 'NB', 'NL', 'NS', 'NT', 'NU',
        'PE', 'QC', 'SK', 'YT', 'AL', 'AK', 'AZ', 'AR', 'CA',
        'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN',
        'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI',
        'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM',
        'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI',
        'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV',
        'WI', 'WY', 'Jan', 'Feb', 'Mar', 'Apr', 'May',
        'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'MM',
        'DD', 'YYYY', "MMM"
    }

    # Default inclusion list
    DEFAULT_INCLUDE = {'fl oz'}

    # Load exclusion and inclusion lists
    exclude = load_custom_list(exclude_filepath, DEFAULT_EXCLUDE)
    include = load_custom_list(include_filepath, DEFAULT_INCLUDE)

    # Predefined patterns (ensure these pattern definitions follow)
    multiple_periods_pattern = re.compile(r'.*\..*\..*')
    postal_code_pattern = re.compile(r'\b[A-Z]\d[A-Z] \d[A-Z]\d\b')

    # Pattern for likely acronyms:
    # - At least two characters.
    # - Primarily uppercase Latin or Greek letters, and numbers.
    # - Can contain hyphens but not start or end with them.
    # - Not solely numbers or hyphens
    # (partially covered by all_digits_or_hyphens_pattern but reinforced).
    likely_pattern = re.compile(
        r'^(?:[A-Z\u0391-\u03A90-9][A-Z\u0391-\u03A90-9-]*[A-Z\u0391-\u03A90-9]|[A-Z\u0391-\u03A90-9]{2,})$')

    # Pattern for possible acronyms:
    # - Words containing at least one sequence of two or more uppercase Latin or Greek letters.
    # - Can be mixed case (e.g., jsHTML, ReqEx).
    possible_pattern = re.compile(
        r'^(?:[A-Za-z0-9\u03B1-\u03C9]*[A-Z\u0391-\u03A9]{2,}[A-Za-z0-9\u03B1-\u03C9]*)+$')

    # Pattern for unlikely acronyms:
    # - Words that start with a capital letter followed by lowercase/numbers,
    #   and may have subsequent capitals followed by lowercase/numbers (e.g., "CamelCase", "Titlecase").
    # - These are often regular capitalized words; is_english_word check is important.
    unlikely_pattern = re.compile(
        r'^[A-Z\u0391-\u03A9][a-z\u03B1-\u03C90-9]+(?:[A-Z\u0391-\u03A9][a-z\u03B1-\u03C90-9]+)*$')

    consecutive_numbers_pattern = re.compile(r'\d{3,}')
    all_digits_or_hyphens_pattern = re.compile(r'^[\d-]+$')
    doc = word_app.ActiveDocument
    text = ' '.join([p.Range.Text for p in doc.Paragraphs])
    text = postal_code_pattern.sub('', text)
    word_list = re.findall(r'\b[\w-]+\b', text)

    acronyms = {'likely': {}, 'possible': {}, 'unlikely': {}}
    prev_word_acronym_index = None

    for i, word in enumerate(word_list):
        if prev_word_acronym_index\
         is not None and i == prev_word_acronym_index + 1:
            prev_word_acronym_index = None
            continue
        elif i < len(word_list) - 1 and re.match(r'^[A-Z]+\d+$', word) and\
                re.match(r'^\d+[A-Z]+$', word_list[i + 1]):
            prev_word_acronym_index = i
            continue

        if (word in exclude or
                consecutive_numbers_pattern.search(word) or
                all_digits_or_hyphens_pattern.match(word) or
                multiple_periods_pattern.match(word)):
            continue

# Core categorization logic:
        # Priority 1: Defined acronyms (and their common variations)
        if word[-1].isdigit() and word[:-1] in defined_acronyms:
            acronyms.setdefault('likely', {})[word[:-1]] = get_context(i, word_list, context_range)
        elif word[-1].lower() == "s" and word[:-1] in defined_acronyms:
            acronyms.setdefault('likely', {})[word[:-1]] = get_context(i, word_list, context_range)
        elif word in defined_acronyms:
            acronyms.setdefault('likely', {})[word] = get_context(i, word_list, context_range)
        
        # Priority 2: Words not in defined_acronyms, evaluated by patterns and English word check.
        elif likely_pattern.match(word): # Matches all-caps or similar acronym structure
            if len(word) <= 3:
                # Short words (e.g., "HDL", "CBS") matching likely_pattern are usually acronyms.
                # Classify as 'likely' even if their lowercase form might be an English word.
                acronyms.setdefault('likely', {})[word] = get_context(i, word_list, context_range)
            elif not is_english_word(word.lower()):
                # Longer words (e.g., "TEAE") matching likely_pattern and NOT English words.
                acronyms.setdefault('likely', {})[word] = get_context(i, word_list, context_range)
            else:
                # Longer words (e.g., "AGAINST", "INTRODUCTION") matching likely_pattern AND ARE English words.
                # These often come from headings or emphasized text.
                acronyms.setdefault('unlikely', {})[word] = get_context(i, word_list, context_range)
        
        elif possible_pattern.match(word) and not is_english_word(word.lower()):
            # Mixed-case with prominent caps (e.g., "ReactComponent", "jsHTML") AND is NOT an English word.
            acronyms.setdefault('possible', {})[word] = get_context(i, word_list, context_range)
        
            # Note: The following block has been removed based on your request:
            # elif unlikely_pattern.match(word) and not is_english_word(word.lower()):
            #     acronyms.setdefault('unlikely', {})[word] = get_context(i, word_list, context_range)
            
            # Note: If a word matches no condition above, it's simply skipped.
        if len(word) in {1, 2} and word.isupper() and\
                i > 0 and i < len(word_list) - 1:
            if word_list[i - 1][0].isupper() and word_list[i + 1][0].isupper():
                for category in acronyms:
                    if word in acronyms[category]:
                        del acronyms[category][word]

    for acronym in list(acronyms['likely']):
        if get_definition(acronym, base_definition_path, user_definition_path) == "":
            # This part of your logic remains if you want to move undefined likely ones to possible
            if acronym in acronyms.get('likely', {}): # Ensure it still exists in likely
                acronyms.setdefault('possible', {})[acronym] = acronyms['likely'][acronym]
                del acronyms['likely'][acronym]

# ... (this comes after the main loop that iterates through word_list
    #      and also after the loop that moves likely acronyms without definitions to possible)

    for phrase_to_include in include:
        # Case-insensitive check for the presence of the phrase in the text
        if phrase_to_include.lower() in text.lower():
            context_value = "Context not available" # Default context
            try:
                # Attempt to find the first occurrence for context (case-insensitive)
                # This provides an estimated word index for get_context.
                char_index = text.lower().index(phrase_to_include.lower())
                # Estimate word index: count words before the char_index.
                # This is an approximation and works best for phrases starting on word boundaries.
                estimated_word_index = len(re.findall(r'\b\w+\b', text[:char_index]))
                context_value = get_context(estimated_word_index, word_list, context_range)
            except ValueError:
                # This can happen if text.lower().index fails, though 'in' check passed (e.g. complex unicode)
                print(f"Warning: Phrase '{phrase_to_include}' found by 'in' but not by index().")
            except Exception as e:
                print(f"Warning: Error getting context for included phrase '{phrase_to_include}': {e}")

            # Categorize the included phrase
            if phrase_to_include in defined_acronyms:
                # If defined, it's 'likely'. Remove from other categories if miscategorized.
                if phrase_to_include in acronyms.get('possible', {}): del acronyms['possible'][phrase_to_include]
                if phrase_to_include in acronyms.get('unlikely', {}): del acronyms['unlikely'][phrase_to_include]
                acronyms.setdefault('likely', {})[phrase_to_include] = context_value
            else:
                # If not defined but in 'include' list, it's considered 'possible'.
                # Ensure it's not already 'likely' (e.g. from a previous broader match).
                if phrase_to_include not in acronyms.get('likely', {}):
                    if phrase_to_include in acronyms.get('unlikely', {}): del acronyms['unlikely'][phrase_to_include]
                    acronyms.setdefault('possible', {})[phrase_to_include] = context_value

    return {
        'likely': {key: acronyms.get('likely', {}).get(key) for key in sorted(acronyms.get('likely', {}))},
        'possible': {key: acronyms.get('possible', {}).get(key) for key in sorted(acronyms.get('possible', {}))},
        'unlikely': {key: acronyms.get('unlikely', {}).get(key) for key in sorted(acronyms.get('unlikely', {}))},
    }


# Other helper functions like get_context,
# get_definition, is_english_word remain the same
def get_context(index, word_list, range):
    start = max(0, index - range)
    end = min(len(word_list), index + range + 1)
    # Adjust end to include the word at the current index
    return ' '.join(word_list[start:end])


def get_definition(acronym, base_definition_path, user_definition_path):
    # Check user definitions first
    if os.path.exists(user_definition_path):
        try:
            with open(user_definition_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if '\t' in line:
                        acr, definition = line.split("\t", 1)
                        if acr == acronym:
                            return definition
        except Exception as e:
            print(f"Warning: Could not read user definitions from {user_definition_path}: {e}")

    # If not found in user definitions, check base definitions
    if os.path.exists(base_definition_path):
        try:
            with open(base_definition_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if '\t' in line:
                        acr, definition = line.split("\t", 1)
                        if acr == acronym:
                            return definition
        except Exception as e:
            print(f"Warning: Could not read base definitions from {base_definition_path}: {e}")

    return "" # Return empty if not found in either


def is_english_word(word):
    if wordnet.synsets(word):
        return True
    else:
        return False


def load_custom_list(filepath, default_set):
    """Loads a custom list from a file, one item per line.
       Items starting with '#' are treated as comments and ignored.
       Returns the custom set if file exists and is readable,
       otherwise returns a copy of the default set.
    """
    custom_set = set()
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    item = line.strip()
                    if item and not item.startswith('#'):
                        custom_set.add(item)
            return custom_set  # Return the loaded custom set
        except Exception as e:
            print(f"Warning: Could not load custom list from {filepath}: {e}")
            return default_set.copy() # Fallback to a copy of the default on error
    return default_set.copy() # Fallback to a copy of the default if file doesn't exist


def _load_definitions_from_file(filepath):
    definitions = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if '\t' in line:
                        acronym, definition = line.split("\t", 1)
                        definitions[acronym] = definition
        except Exception as e:
            print(f"Warning: Could not load definitions from {filepath}: {e}")
    return definitions