import pandas as pd
from bs4 import BeautifulSoup, NavigableString
import pypandoc
import os

# Get the current working directory
cwd = os.path.dirname(__file__)

# Step 1: Convert the .docx file to .html using pypandoc
docx_file = os.path.join(cwd, 'file.docx')
html_file = os.path.join(cwd, 'file.html')
output = pypandoc.convert_file(docx_file, 'html', outputfile=html_file)
assert output == ""

# Step 2: Read the .csv file and create the dictionary
csv_file = os.path.join(cwd, 'abbr.csv')
df = pd.read_csv(csv_file)
abbr_dict = df.set_index('Acronym')['Title'].to_dict()

# Step 3: Parse the HTML
with open(html_file, 'r') as f:
    soup = BeautifulSoup(f, 'html.parser')

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
            if acronym in text_node:
                new_content = text_node.replace(acronym, f'<abbr title="{title}">{acronym}</abbr>')
                text_node.replace_with(BeautifulSoup(new_content, 'html.parser'))

# Step 5: Add a space before every <abbr> tag if there is a leading space in the <abbr> tag
for abbr in soup.find_all('abbr'):
    if abbr.string and abbr.string[0] == ' ':
        abbr.string = abbr.string[1:]
        abbr.insert_before(' ')

# Step 6: Write the modified HTML back to a file
output_file = os.path.join(cwd, 'output.html')
with open(output_file, 'w') as f:
    f.write(str(soup))