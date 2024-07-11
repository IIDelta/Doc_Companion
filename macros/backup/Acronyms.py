from nltk.corpus import wordnet
import re

def find_acronyms(word_app, definition_path, context_range=5):
    # Load the acronyms from the definition file
    with open(definition_path, 'r') as f:
        defined_acronyms = {line.split('\t')[0] for line in f.read().splitlines()}

    doc = word_app.ActiveDocument
    text = ' '.join([p.Range.Text for p in doc.Paragraphs])

    # Add a condition to exclude words with more than one '-'
    hyphenated_pattern = re.compile(r'.*-.*-.*')

    # Add a condition to exclude words with more than one '.'
    multiple_periods_pattern = re.compile(r'.*\..*\..*')


    # Remove postal codes from the text
    postal_code_pattern = re.compile(r'\b[A-Z]\d[A-Z] \d[A-Z]\d\b')
    text = postal_code_pattern.sub('', text)

    # Split the text into words, treating hyphens as part of the words
    word_list = re.findall(r'\b[\w-]+\b', text)

    # Define the regex patterns for the three categories
    # Add unicode ranges for Greek letters
    likely_pattern = re.compile(r'^[\u0391-\u03A9A-Z0-9-]{2,}$')
    possible_pattern = re.compile(r'^.*[\u0391-\u03A9A-Z]{2,}.*$')
    unlikely_pattern = re.compile(r'^.*[\u0391-\u03A9A-Z]{2,}.*$')


    # Exclude specific acronyms, postal codes, abbreviations for provinces, states, and months
    exclude = {
        'ON', 'BC', 'MB', 'NB', 'NL', 'NS', 'NT', 'NU', 'PE', 'QC', 'SK', 'YT',  # Canadian provinces
        'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA', 'KS',
        'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM', 'NY',
        'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV',
        'WI', 'WY',  # American states
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',  # Months
        'MM', 'DD', 'YYYY', "MMM" # dates placeholders
    }
    
    include = {
        'fl oz', # phrases (spaces separation)
    }
    
    consecutive_numbers_pattern = re.compile(r'\d{3,}')
    all_digits_or_hyphens_pattern = re.compile(r'^[\d-]+$')  # Matches words that are made up entirely of digits and/or hyphens
        
    acronyms = {'likely': {}, 'possible': {}, 'unlikely': {}}
    
    prev_word_acronym = False
    prev_word_acronym_index = None
    for i, word in enumerate(word_list):
        # If the previous word was an acronym, skip this word
        if prev_word_acronym:
            prev_word_acronym = False
            continue
        elif i < len(word_list) - 1 and re.match(r'^[A-Z]+\d+$', word) and re.match(r'^\d+[A-Z]+$', word_list[i + 1]):
            # The current word and the next one together look like a postal code, skip them
            prev_word_acronym = True
            # Save the index of the previous acronym
            prev_word_acronym_index = i
            continue
        elif prev_word_acronym_index is not None and i == prev_word_acronym_index + 1:
            # Skip the previous acronym if it was part of a postal code
            prev_word_acronym_index = None
            continue
    
    prev_word_vitamin = False  # Track if the previous word is 'vitamin'
    for i, word in enumerate(word_list):
        if (word in exclude or consecutive_numbers_pattern.search(word) or 
            all_digits_or_hyphens_pattern.match(word) or hyphenated_pattern.match(word) or
            multiple_periods_pattern.match(word)):
            prev_word_acronym = False
            continue

        # If the previous word was 'vitamin', skip this word
        if prev_word_vitamin:
            prev_word_vitamin = False
            continue

        # If the previous word was an acronym, skip this word
        if prev_word_acronym:
            prev_word_acronym = False
            continue

        
        # Check if the last character of the acronym is a digit and if the acronym without the last character is in the defined acronyms
        if word[-1].isdigit() and word[:-1] in defined_acronyms:
            acronyms['likely'][word[:-1]] = get_context(i, word_list, context_range)
            prev_word_acronym = True
            continue

        # Check if the last character of the acronym is "s" and if the acronym without the last character is in the defined acronyms
        elif word[-1].lower() == "s" and word[:-1] in defined_acronyms:
            acronyms['likely'][word[:-1]] = get_context(i, word_list, context_range)
            prev_word_acronym = True
            continue

        # Check for likely acronyms
        if word in defined_acronyms:
            acronyms['likely'][word] = get_context(i, word_list, context_range)
            prev_word_acronym = True
        elif likely_pattern.match(word) and not is_english_word(word.lower()):
            acronyms['likely'][word] = get_context(i, word_list, context_range)
            prev_word_acronym = True
        elif possible_pattern.match(word) and not is_english_word(word.lower()):
            acronyms['possible'][word] = get_context(i, word_list, context_range)
            prev_word_acronym = True
        elif word.lower() == 'vitamin':
            prev_word_vitamin = True
        elif unlikely_pattern.match(word) and word not in defined_acronyms:
            acronyms['unlikely'][word] = get_context(i, word_list, context_range)
            prev_word_acronym = True
        else:
            prev_word_acronym = False

        # If the acronym is two capital letters, check if the words before and after are capitalized
        if len(word) in {1, 2} and word.isupper() and i > 0 and i < len(word_list) - 1:
            if word_list[i - 1][0].isupper() and word_list[i + 1][0].isupper():
                # If the words before and after are capitalized, it's likely an initial, so remove it from the acronyms
                for category in acronyms:
                    if word in acronyms[category]:
                        del acronyms[category][word]

    # After populating the acronyms dictionary, check if the 'likely' acronyms have definitions
    for acronym in list(acronyms['likely']):
        if get_definition(acronym, definition_path) == "":
            # If an acronym doesn't have a definition, move it to 'possible'
            acronyms['possible'][acronym] = acronyms['likely'][acronym]
            del acronyms['likely'][acronym]

    # Add the abbreviations from the 'include' variable to the 'likely' category if found in the word document
    for phrase in include:
        if phrase in text and phrase in defined_acronyms:
            acronyms['likely'][phrase] = get_context(text.index(phrase), word_list, context_range)
    
    return {
    'likely': {key: acronyms['likely'][key] for key in sorted(acronyms['likely'])},
    'possible': {key: acronyms['possible'][key] for key in sorted(acronyms['possible'])},
    'unlikely': {key: acronyms['unlikely'][key] for key in sorted(acronyms['unlikely'])},
    }         


def get_context(index, word_list, range):
    start = max(0, index - range)
    end = min(len(word_list), index + range + 1)  # Adjust end to include the word at the current index
    return ' '.join(word_list[start:end])

def get_definition(acronym, definition_path):
    definitions = {}
    with open(definition_path, "r") as f:
        for line in f:
            acronym_definition, definition = line.strip().split("\t")
            definitions[acronym_definition] = definition
    return definitions.get(acronym, "")

def is_english_word(word):
    if wordnet.synsets(word):
        return True
    else:
        return False

