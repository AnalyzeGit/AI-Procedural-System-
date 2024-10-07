import spacy
from typing import List, Tuple, Optional

# spaCy model
nlp = spacy.load('en_core_web_sm')

# Define the return types for clarity
PosTag = Tuple[str, str]
ParseTree = Tuple[str, str]
NamedEntity = Tuple[str, str]
ExtractionResult = Tuple[Optional[str], Optional[str], Optional[str]]


def process_text(text: str) -> Tuple[List[PosTag], List[ParseTree], List[NamedEntity]]:
    """
    Processes the given text using spaCy to extract part-of-speech tags,
    parse tree elements, and named entities.

    :param text: The text to be processed.
    :type text: str
    :return: A tuple containing lists of POS tags, parse tree elements,
             and named entities.
    :rtype: Tuple[List[PosTag], List[ParseTree], List[NamedEntity]]
    """
    doc = nlp(text)
    pos_tags = [(token.text, token.tag_) for token in doc]
    parse_tree = [(token.text, token.dep_) for token in doc]
    ner = [(ent.text, ent.label_) for ent in doc.ents]
    return pos_tags, parse_tree, ner


def extract_verb_object(sentence: str) -> ExtractionResult:
    """
        Extracts the main verb and its direct or prepositional object
        from the given sentence. Identifies the root verb and returns its
        details if it's a verb.

        :param sentence: The sentence from which to extract the verb-object pair.
        :type sentence: str
        :return: A tuple containing the paragraph type, verb, and object (if any).
        :rtype: ExtractionResult
    """
    doc = nlp(sentence)
    para_type, verb, obj = None, None, None

    for token in doc:
        if token.dep_ == 'ROOT' and token.tag_ == 'VB':
            verb = token.text
            for child in token.children:
                if child.dep_ in ['dobj', 'pobj']:  # direct or prepositional object
                    obj = child.text
                    para_type = 'VS'
                    break
    return para_type, verb, obj


def extract_verb_object_after_then(sentence: str) -> ExtractionResult:
    """
        Splits the sentence at 'THEN' and processes the part after it to extract
        the main verb and its object. Specifically designed to handle sentences
        containing the word 'THEN'.

        :param sentence: The sentence to process.
        :type sentence: str
        :return: A tuple of paragraph type, verb, and object extracted
                 from the part of the sentence after 'THEN'.
        :rtype: ExtractionResult
    """
    parts = sentence.split('THEN')
    if len(parts) > 1:
        after_then = parts[1].strip()  # Sentence part after 'THEN'
        doc = nlp(after_then)
        para_type, verb, obj = None, None, None

        for token in doc:
            if token.dep_ == 'ROOT' and token.tag_ == 'VB':
                verb = token.text
                for child in token.children:
                    if child.dep_ in ['dobj', 'pobj']:  # direct or prepositional object
                        obj = child.text
                        para_type = 'CndnAS'
                        break
        return para_type, verb, obj
    return None, None


def extract_verb_object_from_when(sentence: str) -> ExtractionResult:
    """
        Analyzes sentences that start with 'WHEN', splitting at commas to find
        the imperative part, and then extracts the verb-object pair from it.

        :param sentence: The sentence to analyze.
        :type sentence: str
        :return: A tuple of paragraph type, verb, and object extracted from the
                 imperative part of the sentence.
        :rtype: ExtractionResult
    """
    # Splitting at commas to find the imperative part
    parts = sentence.split(',')
    if len(parts) > 1:
        imperative_part = parts[-1].strip()  # Consider the last part after splitting
        doc = nlp(imperative_part)
        para_type, verb, obj = None, None, None

        for token in doc:
            # Check if the token is a verb (imperative verbs are usually at the start)
            if token.head == token and token.pos_ == 'VERB':
                verb = token.text
                for child in token.children:
                    # Finding the object of the verb
                    if child.dep_ in ['dobj', 'pobj']:  # direct or prepositional object
                        obj = child.text
                        para_type = 'CndnAS'
                        break
        return para_type, verb, obj
    return None, None, None


def extract_verb_object_from_whenever(sentence: str) -> ExtractionResult:
    """
        Targets sentences starting with 'WHILE' or 'WHENEVER', splits them
        at commas, and processes the last part to extract the verb-object pair.

        :param sentence: The sentence to analyze.
        :type sentence: str
        :return: A tuple of paragraph type, verb, and object from the last part
                 of the sentence.
        :rtype: ExtractionResult
    """
    # Splitting at commas to find the imperative part
    parts = sentence.split(',')
    if len(parts) > 1:
        imperative_part = parts[-1].strip()  # Consider the last part after splitting
        doc = nlp(imperative_part)
        para_type, verb, obj = None, None, None

        for token in doc:
            # Check if the token is a verb (imperative verbs are usually at the start)
            if token.head == token and token.pos_ == 'VERB':
                verb = token.text
                for child in token.children:
                    # Finding the object of the verb
                    if child.dep_ in ['dobj', 'pobj']:  # direct or prepositional object
                        obj = child.text
                        para_type = 'CntsAS'
                        break
        return para_type, verb, obj
    return None, None, None


def combined_extraction(sentence: str) -> ExtractionResult:
    """
        Analyzes a sentence and extracts different types of information based on
        specific patterns or keywords. This function looks for particular patterns
        such as the presence of specific words ('signature', 'date', 'THEN', etc.)
        or symbols ('_', '=') and, based on these, calls the appropriate extraction
        function or returns a predefined paragraph type.

        :param sentence: The sentence to be analyzed and processed.
        :type sentence: str
        :return: A tuple containing the paragraph type and, optionally, the main verb
                 and its object. The paragraph type can be 'Signoff', 'CalcRow', or
                 types identified by other specific extraction functions.
        :rtype: ExtractionResult

        The function applies different rules:
        - If the sentence contains '_' and 'signature' or 'date', it is classified as 'Signoff'.
        - If the sentence contains both '=' and '_', it is classified as 'CalcRow'.
        - If the sentence starts with 'WHILE' or 'WHENEVER', it calls `extract_verb_object_from_whenever`.
        - If the sentence starts with 'WHEN', it calls `extract_verb_object_from_when`.
        - If 'THEN' is in the sentence, it calls `extract_verb_object_after_then`.
        - Otherwise, it defaults to calling `extract_verb_object`.
    """
    if '_' in sentence and ('signature' in sentence.lower() or 'date' in sentence.lower()):
        return 'Signoff', None, None
    elif '=' in sentence and '_' in sentence:
        return 'CalcRow’', None, None
    elif sentence.upper().startswith('WHILE') or sentence.upper().startswith('WHENEVER'):
        return extract_verb_object_from_whenever(sentence)
    elif sentence.upper().startswith('WHEN'):
        return extract_verb_object_from_when(sentence)
    elif 'THEN' in sentence.upper():
        return extract_verb_object_after_then(sentence)
    else:
        return extract_verb_object(sentence)