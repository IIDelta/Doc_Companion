from nltk.corpus import wordnet
import re


def find_acronyms(word_app, definition_path, context_range=5):
    # Predefined patterns
    hyphenated_pattern = re.compile(r'.*-.*-.*')
    multiple_periods_pattern = re.compile(r'.*\..*\..*')
    postal_code_pattern = re.compile(r'\b[A-Z]\d[A-Z] \d[A-Z]\d\b')
    likely_pattern = re.compile(r'^[\u0391-\u03A9A-Z0-9-]{2,}$')
    possible_pattern = re.compile(r'^.*[\u0391-\u03A9A-Z]{2,}.*$')
    unlikely_pattern = re.compile(r'^.*[\u0391-\u03A9A-Z]{2,}.*$')
    consecutive_numbers_pattern = re.compile(r'\d{3,}')
    all_digits_or_hyphens_pattern = re.compile(r'^[\d-]+$')
    # Exclude and include lists
    exclude = {
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
    include = {'fl oz'}

    with open(definition_path, 'r') as f:
        defined_acronyms = {
            line.split('\t')[0] for line in f.read().splitlines()}

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

        if word in exclude or consecutive_numbers_pattern.search(word) or \
            all_digits_or_hyphens_pattern.match(word) or\
            hyphenated_pattern.match(word) or\
                multiple_periods_pattern.match(word):
            continue

        if word[-1].isdigit() and word[:-1] in defined_acronyms:
            acronyms['likely'][word[:-1]] = get_context(
                i, word_list, context_range)
        elif word[-1].lower() == "s" and word[:-1] in defined_acronyms:
            acronyms['likely'][word[:-1]] = get_context(
                i, word_list, context_range)
        elif word in defined_acronyms:
            acronyms['likely'][word] = get_context(
                i, word_list, context_range)
        elif likely_pattern.match(word) and not\
                is_english_word(word.lower()):
            acronyms['likely'][word] = get_context(
                i, word_list, context_range)
        elif possible_pattern.match(word) and not\
                is_english_word(word.lower()):
            acronyms['possible'][word] = get_context(
                i, word_list, context_range)
        elif word.lower() == 'vitamin':
            continue
        elif unlikely_pattern.match(word) and word not in defined_acronyms:
            acronyms['unlikely'][word] = get_context(
                i, word_list, context_range)

        if len(word) in {1, 2} and word.isupper() and\
                i > 0 and i < len(word_list) - 1:
            if word_list[i - 1][0].isupper() and word_list[i + 1][0].isupper():
                for category in acronyms:
                    if word in acronyms[category]:
                        del acronyms[category][word]

    for acronym in list(acronyms['likely']):
        if get_definition(acronym, definition_path) == "":
            acronyms['possible'][acronym] = acronyms['likely'][acronym]
            del acronyms['likely'][acronym]

    for phrase in include:
        if phrase in text and phrase in defined_acronyms:
            acronyms['likely'][phrase] = get_context(
                text.index(phrase), word_list, context_range)

    return {
        'likely': {key: acronyms['likely'][key] for key in sorted(
            acronyms['likely'])},
        'possible': {key: acronyms['possible'][key] for key in sorted(
            acronyms['possible'])},
        'unlikely': {key: acronyms['unlikely'][key] for key in sorted(
            acronyms['unlikely'])},
    }


# Other helper functions like get_context,
# get_definition, is_english_word remain the same
def get_context(index, word_list, range):
    start = max(0, index - range)
    end = min(len(word_list), index + range + 1)
    # Adjust end to include the word at the current index
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
