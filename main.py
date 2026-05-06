import pandas as pd
import string
import re

excel_file="Papers_Authors.xlsx"
abstract_file="text_with_abstracts.txt"

df=pd.read_excel(excel_file)

def clean_title(text):
    title=text.strip().title()
    title = re.sub(r"\s*\n\s*", " ", title)
    title = re.sub(r"\s{2,}", " ", title)
    return title

def to_superscript(text):
    sup_map={
        'a': 'ᵃ','b': 'ᵇ','c': 'ᶜ','d': 'ᵈ','e': 'ᵉ',
        'f': 'ᶠ','g': 'ᵍ','h': 'ʰ','i': 'ᶦ','j': 'ʲ',
        'k': 'ᵏ','l': 'ˡ','m': 'ᵐ','n': 'ⁿ','o': 'ᵒ',
        'p': 'ᵖ','q': 'ᑫ','r': 'ʳ','s': 'ˢ','t': 'ᵗ',
        'u': 'ᵘ','v': 'ᵛ','w': 'ʷ','x': 'ˣ','y': 'ʸ',
        'z': 'ᶻ'
    }
    return ''.join(sup_map.get(char,char) for char in text)

def parse_abstracts(file_path):
    data = {}

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    # split submissions
    submissions = re.split(r"Submission\s+#?(\d+):", content)

    for i in range(1, len(submissions), 2):
        paper_id = int(submissions[i])
        block = submissions[i + 1].strip()

                    #  TITLE 
        title_match = re.match(r"^(.*?)(?=\n=+)", block, re.DOTALL)
        title = ""

        if title_match:
            first_line = title_match.group(1).strip()
            # remove any trailing newlines and normalize spacing
            title = clean_title(first_line)

                    #  ABSTRACT 
        abstract_match = re.search(
            r"Abstract\s*\n[-–—]+\n(.*?)(?=\n(?:Author|Authors|Presenter|Keywords|Topic|Topics|Submission)\b)",
            block,
            re.DOTALL | re.IGNORECASE
        )
        abstract = abstract_match.group(1).strip() if abstract_match else ""

                    #  KEYWORDS 
        keywords_match = re.search(
            r"Keywords\s*\n[-–—]+\n(.*?)(?=\n(?:Topic|Topics|Author|Authors|Presenter|Submission)\b)",
            block,
            re.DOTALL | re.IGNORECASE
        )
        keywords = keywords_match.group(1).strip() if keywords_match else ""

        abstract = re.sub(r"\n{2,}", "\n\n", abstract)
        keywords = re.sub(r"\n+", " ", keywords)

        data[paper_id] = {
            "title": title,
            "abstract": abstract,
            "keywords": keywords
        }

    return data

def format_authors(group):
    affiliations = group['affiliation'].unique()
    aff_map = {aff: string.ascii_lowercase[i] for i, aff in enumerate(affiliations)}
    
    authors_list = []
    emails_list = []
    
    for _, row in group.iterrows():
        first = row['given_name'].capitalize()
        last = row['family_name'].upper()
        
        aff_letter = aff_map[row['affiliation']]
        
        is_contact = str(row.get('iscontact', '')).strip().lower() == "yes"
        contact_mark = "*" if is_contact else ""
        
        sup = to_superscript(aff_letter)
        
        author_str = f"{first} {last}{sup}{contact_mark}"
        authors_list.append(author_str)
        
        email = str(row.get('email', '')).strip()
        email_str = f"(*) {email}" if is_contact else email
        emails_list.append(email_str)
    
    return ",".join(authors_list), aff_map, "; ".join(emails_list)

final_rows=[]

for paper_id,group in df.groupby("paper"):
    authors_str, aff_map, emails_str = format_authors(group)
    
    aff_strings = []
    for aff, letter in aff_map.items():
        sup_letter = to_superscript(letter)
        aff_strings.append(f"{sup_letter} {aff}")

    affiliations_str = "\n".join(aff_strings)
    
    abstracts_dict = parse_abstracts(abstract_file)
    
    paper_data = abstracts_dict.get(paper_id, {})
    abstract = paper_data.get("abstract", "")
    keywords = paper_data.get("keywords", "")
    title = paper_data.get("title", "")
    
    final_rows.append({
        "paper ID": paper_id,
        "title": title,
        "authors": authors_str,
        "affiliations": affiliations_str,
        "emails": emails_str,
        "abstract": abstract,
        "keywords": keywords
    })

final_df=pd.DataFrame(final_rows)

output_file="final_papers.xlsx"
final_df.to_excel(output_file,index=False)

print(f"Job done: {output_file}")