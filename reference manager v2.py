import re
import docx
from habanero import Crossref
from datetime import datetime

# --- CONFIGURATION ---
# Your email address
EMAIL_ADDRESS = "example@gmail.com" 

# Path to the input and output Word files.
INPUT_FILE = r"C:\Users\PC\Desktop\Filename.docx"
OUTPUT_FILE = r"C:\Users\PC\Desktop\Edited File.docx" #It will save the editted file in the path you write here

# Create an instance of the Crossref client.
cr = Crossref(mailto=EMAIL_ADDRESS)

def format_authors(authors_list):
    """Formats a list of author objects into 'J.D. Doe, F. Bar' style."""
    if not authors_list:
        # The reference might have an organization name instead of an author.
        return ""
    
    formatted_authors = []
    for author in authors_list:
        # If a family name exists.
        if 'family' in author:
            given_name = author.get('given', '')
            # Create initials from the given name.
            initials = '.'.join([name[0] for name in given_name.split()]) + '.' if given_name else ''
            formatted_authors.append(f"{initials} {author['family']}")
        # If only a general name exists (common for organizations).
        elif 'name' in author:
            formatted_authors.append(author['name'])
            
    return ', '.join(formatted_authors)

def format_journal_article(item):
    """Formats a reference for a journal article."""
    authors = format_authors(item.get('author'))
    title = '. '.join(item.get('title', ['']))
    journal = '. '.join(item.get('container-title', ['']))
    volume = item.get('volume', '')
    pages = item.get('page', '')
    article_number = item.get('article-number', '')
    year = item.get('issued', {}).get('date-parts', [[None]])[0][0]
    doi = item.get('DOI', '')

    # Build the final string.
    ref = f"{authors}, {title}, {journal}"
    if volume:
        ref += f" {volume}"
    if year:
        ref += f" ({year})"
    if pages:
        ref += f" {pages.replace('-', '–')}" # Use en-dash
    elif article_number: # For articles that have an article number instead of pages.
        ref += f", {article_number}"
    
    ref += "."
    if doi:
        ref += f" https://doi.org/{doi}."
        
    return ref

def format_book(item):
    """Formats a reference for a book."""
    authors = format_authors(item.get('author'))
    title = '. '.join(item.get('title', ['']))
    edition = item.get('edition-number', '')
    publisher = item.get('publisher', '')
    publisher_place = item.get('publisher-place', '')
    year = item.get('issued', {}).get('date-parts', [[None]])[0][0]
    
    ref = f"{authors}, {title}"
    if edition:
        # Converting number to ordinal string (e.g., 4 -> "fourth") is complex; using a simpler format.
        ref += f", {edition}. ed."
    if publisher:
        ref += f", {publisher}"
    if publisher_place:
        ref += f", {publisher_place}"
    if year:
        ref += f", {year}."
        
    return ref.replace('..', '.').replace(',,', ',')

def format_book_chapter(item):
    """Formats a reference for a book chapter."""
    authors = format_authors(item.get('author'))
    editors = format_authors(item.get('editor'))
    chapter_title = '. '.join(item.get('title', ['']))
    book_title = '. '.join(item.get('container-title', ['']))
    publisher = item.get('publisher', '')
    publisher_place = item.get('publisher-place', '')
    year = item.get('issued', {}).get('date-parts', [[None]])[0][0]
    pages = item.get('page', '')
    
    ref = f"{authors}, {chapter_title}"
    if editors:
        ref += f", in: {editors} (Eds.), {book_title}"
    else:
        ref += f", in: {book_title}"
    
    if publisher:
        ref += f", {publisher}"
    if publisher_place:
        ref += f", {publisher_place}"
    if year:
        ref += f", {year}"
    if pages:
        ref += f", pp. {pages.replace('-', '–')}."
    else:
        ref += "."
        
    return ref

def format_dataset(item):
    """Formats a reference for a dataset."""
    authors = format_authors(item.get('author'))
    title = '. '.join(item.get('title', ['']))
    publisher = item.get('publisher', '')
    version = item.get('version', '')
    year = item.get('issued', {}).get('date-parts', [[None]])[0][0]
    doi = item.get('DOI', '')

    ref = f"{authors}, {title} [dataset], {publisher}"
    if version:
        ref += f", v{version}"
    if year:
        ref += f", {year}."
    if doi:
        ref += f" https://doi.org/{doi}."
        
    return ref

def format_web_reference(item):
    """Formats a reference for a web resource (best-effort)."""
    # Crossref often lacks complete metadata for websites, so this function makes a best effort.
    authors = format_authors(item.get('author', []) or [{'name': item.get('publisher', '')}])
    title = '. '.join(item.get('title', ['']))
    url = item.get('URL', '')
    year = item.get('issued', {}).get('date-parts', [[None]])[0][0]
    today = datetime.now().strftime("%d %B %Y")

    ref = f"{authors}, {title}. {url}"
    if year:
        ref += f", {year}"
    
    ref += f" (accessed {today})."
    return ref


def process_references():
    """Opens the Word document, processes references, and saves the new file."""
    try:
        doc = docx.Document(INPUT_FILE)
    except Exception as e:
        print(f"Error opening document: {e}")
        return

    report = []
    print("Processing references... This may take a while.")

    for i, p in enumerate(doc.paragraphs):
        # Find paragraphs that start with a number in brackets.
        match = re.match(r'^\s*(\[\d+\])\s*(.*)', p.text)
        if not match:
            continue

        ref_number = match.group(1)
        query_text = match.group(2).strip()
        original_text = p.text.strip()
        
        print(f"Querying for: {query_text[:70]}...")

        try:
            # Query Crossref.
            results = cr.works(query=query_text, limit=1)
            
            if not results['message']['items']:
                report.append(f"NOT FOUND: {ref_number} - Could not find metadata for '{query_text[:50]}...'.")
                continue

            item = results['message']['items'][0]
            ref_type = item.get('type', 'unknown')
            
            new_text_body = ""
            # Select the formatting function based on the reference type.
            if ref_type == 'journal-article':
                new_text_body = format_journal_article(item)
            elif ref_type == 'book':
                new_text_body = format_book(item)
            elif ref_type == 'book-chapter':
                new_text_body = format_book_chapter(item)
            elif ref_type == 'dataset':
                new_text_body = format_dataset(item)
            elif ref_type == 'component' and 'software' in item.get('title', [''])[0].lower(): # Heuristic for software.
                 new_text_body = format_dataset(item).replace('[dataset]', '[software]') # Temporarily use the dataset format.
            else:
                # For other types or websites, use a generic format.
                new_text_body = format_web_reference(item)

            new_full_text = f"{ref_number} {new_text_body}"

            # If the text has changed, replace it.
            if new_full_text.strip() != original_text:
                p.text = new_full_text
                report.append(f"CHANGED: {ref_number} - Successfully reformatted.")
            else:
                report.append(f"UNCHANGED: {ref_number} - Already in correct format or no changes needed.")

        except Exception as e:
            report.append(f"ERROR: {ref_number} - An error occurred while processing: {e}")

    # Save the final document.
    try:
        doc.save(OUTPUT_FILE)
        print(f"\nProcessing complete. Formatted document saved as '{OUTPUT_FILE}'.")
    except Exception as e:
        print(f"Error saving document: {e}")

    # Print the final report.
    print("\n--- Processing Report ---")
    if not report:
        print("No references starting with '[number]' were found to process.")
    for line in report:
        print(line)
    print("-------------------------\n")


if __name__ == "__main__":
    process_references()
