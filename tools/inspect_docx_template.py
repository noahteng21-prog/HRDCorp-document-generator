import os
import zipfile
import re

path = os.path.join('templates', 'New DOCX PSMB_SBL_KHAS_T3_01 template.docx')
print('path:', path)
print('exists:', os.path.exists(path))
with zipfile.ZipFile(path) as zf:
    data = zf.read('word/document.xml').decode('utf-8', errors='replace')
    print('has {{:', '{{' in data)
    print('has {%%:', '{%' in data)
    matches = re.findall(r'\{[{%].*?[}%]\}', data)
    print('matches:', matches[:100])
