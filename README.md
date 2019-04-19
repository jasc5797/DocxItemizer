# DocxItemizer
Python script that itemizes the components of a document (.docx file). 

DocxItemizer is intended to be used in forensic investigations. This script will extract all of the contents of a document into a separate directory for investigation. Multiple documents can be itemized at the same time by passing a path to a directory instead of a path to a document. The resulting files are made viewable in different two ways. The contents in their original file structure can be viewed in the file system, and the contents are also itemized into separate folders depending on their file type.

The user will be alerted if any image files are found in the .docx file that have the wrong file extension. This can be useful for finding image files that a user may have try to hide by using the wrong extension(e.g. ".txt"). If any images are found an additional directory will be created containing the hidden images.

An optional search term can also be provided as an argument when running the script. The search term should be in a regex format. The user will be alerted if any file names or file contents match the regex. If any files match the regex an additional directory will be created containing the matching files.

# Itemization
The contents of the document are itemized into these categories
* **XML**: XML files
* **CSS**: CSS files
* **Media**: Images and other media
* **Content**: Extracted text from the document separated by location (e.g. document, header, footer, etc.)
* **RELS**: RELS files
* **Uncategorized**: Files that do not fit into the previous categories

# Prerequisites
Packages needed to run this script:
* lxml


#  Samples
## Single Document
Itemize a single document
```
python3 docxitemizer.py [path to .docx file]
```

## Multiple Documents
Itemize all documents in a directory
```
python3 docitemizer.py [path to directory containing .docx file(s)]
```

## Optional Regex Search Term
List all files that their name or contents match a regex expression
```
python3 docitemizer.py [path] [search term]
```

## Help 
View help in the command line
```
python3 docitemizer.py -h
```
### Help Output
```
usage: DocxItemizer.py [-h] path [search_term]

Docx Itemizer

positional arguments:
  path         Required Argument: Path to .docx file or directory containing
               .docx file(s)
  search_term  Optional Argument: Regex to use to match file names and file
               contents

optional arguments:
  -h, --help   show this help message and exit
 ```
