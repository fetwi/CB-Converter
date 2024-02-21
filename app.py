import pandas as pd
from bs4 import BeautifulSoup, NavigableString
import pypandoc
import os
import re
import streamlit as st

st.markdown(
    """
<style>
    #MainMenu { display:none;}
</style>
"""
)

st.set_page_config(
    page_title="Convertwi",
    page_icon="favicon.png",
    initial_sidebar_state="expanded",
)

# Get the current working directory
cwd = os.path.dirname(__file__)

# Step 1: Convert the .docx file to .html using pypandoc
docx_file = st.file_uploader("Upload the script file (.docx format).", type="docx")
if docx_file is not None:
    docx_file_path = os.path.join(cwd, 'file.docx')
    with open(docx_file_path, 'wb') as f:
        f.write(docx_file.getbuffer())
    html_file = os.path.join(cwd, 'file.html')
    output = pypandoc.convert_file(docx_file_path, 'html', outputfile=html_file, extra_args=['--no-highlight'])
    assert output == ""

    # Step 2: Read the .csv file and create the dictionary
    # Specify the path to your local CSV file
    csv_file_path = os.path.join(cwd, 'abbr.csv')
    df = pd.read_csv(csv_file_path)
    abbr_dict = df.set_index('Acronym')['Title'].to_dict()

    # Sort the dictionary by length of acronym, longest first
    abbr_dict = {k: v for k, v in sorted(abbr_dict.items(), key=lambda item: -len(item[0]))}

    # Step 3: Parse the HTML
    with open(html_file, 'r') as f:
        soup = BeautifulSoup(f, 'html.parser')

        # Remove all code before the first <h1> tag
        first_h1 = soup.find('h1')
        for tag in list(first_h1.previous_siblings):
            tag.extract()

        # Replace all &nbsp; with a space within <p> tags
        for p in soup.find_all('p'):
            for content in p.contents:
                if isinstance(content, NavigableString):
                    content.replace_with(content.replace(u'\xa0', ' '))

        # Step 4: Iterate over the text nodes and replace the acronyms
        for acronym, title in abbr_dict.items():
            if isinstance(acronym, str) and isinstance(title, str):
                # Collect all the text nodes
                text_nodes = [node for node in soup.find_all(text=True) if isinstance(node, NavigableString)]
                # Create a copy of the text nodes list
                text_nodes_copy = text_nodes.copy()
                for text_node in text_nodes_copy:
                    # Ignore if the acronym is already in an <abbr> tag
                    if text_node.parent.name == 'abbr':
                        continue
                    # Ignore if the acronym is directly followed by a capital character
                    if re.search(f'{acronym}[A-Z]', text_node):
                        continue
                    if acronym in text_node:
                        new_content = text_node.replace(acronym, f'<abbr title="{title}">{acronym}</abbr>')
                        text_node.replace_with(BeautifulSoup(new_content, 'html.parser'))

        # Step 5: Add a space before every <abbr> tag if there is a leading space in the <abbr> tag
        for abbr in soup.find_all('abbr'):
            if abbr.string and abbr.string[0] == ' ':
                abbr.string = abbr.string[1:]
                abbr.insert_before(' ')

    # Step 6: Display the modified HTML in a code block and copy it to the clipboard
    modified_html = str(soup)
    st.code(modified_html, language='html')

    # Remove the temporary files
    os.remove(docx_file_path)
    os.remove(html_file)