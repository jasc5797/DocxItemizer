import zipfile              # Extracts from zip files
import os                   # Interacts with the file system
import shutil               # Makes copies of files
import time                 # Gets a time stamps
import imghdr               # Checks if a file is actually an image
import re                   # Regex search file name and file contents
import lxml.etree as et     # Processes the XML files and transform them using XSLT. This package will need to be installed in order to run the script
import argparse             # Parses the arguments passed by the user. Also provides a help menu using the [-h] flag

__author__ = 'James Stinson-Cerra'
__date__ = '20190417'
__version__ = '1'

# Sets up the parser for the Docx Itemizer script
parser = argparse.ArgumentParser(description="Docx Itemizer")
# Non Optional Argument: Path to a .docx file or directory containing .docx file(s)
parser.add_argument("path", type=str,
                    help="Required Argument: Path to .docx file or directory containing .docx file(s)")
# Optional Argument: Term to search for in file names and file contents while itemizing the .docx file(s)
parser.add_argument("search_term", type=str, nargs="?",
                    help="Optional Arguement: Regex to use to match file names and file contents")


class Itemizer:

    def __init__(self, doc_path, base_dir_path, extracted_dir_path, zip_dir_path, doc_copy_path):
        """
        The Itemizer class is responsible for itemizing the different components found within a .docx file

        1. Extracts the documents contents by converting it to a .zip and unzipping
            it into "Extracted Document" directory
        2. Copies original .docx file to base directory
        3. Itemizes the contents of the "Extracted Document" directory into separate directories:
            XML, CSS, Media, Content, RELS, and Uncatergorized
            a. XML Directory: Contains all XML files found in the "Extracted Document" directory
            b. CSS Directory: Contains all CSS files found in the "Extracted Document" directory
            c. Media Directory: Contains all media files found in the extracted files in the media directory
                * Media directory found at "Extracted Document/{Doc Name}/word/media" or "Extracted Document/word/media"
            d. Content Directory: Contains all XML files from the word directory converted to plain text
                * Word directory found at "Extracted Document/{Doc Name}/word" or "Extracted Document/word"
                * XML contents is converted to plain text and stripped of XML elements
                    by using XSLT and some python to insert newlines
            e. RELS Directory: Contains all RELS files found in the "Extracted Document" directory
            f. Uncatergorized Directory: Contains all files with unknown file extensions
                in the "Extracted Document" directory

        :param doc_path: Path to the document
        :param base_dir_path: Path to the base directory that holds all of the output of this script
        :param extracted_dir_path: Path to the "Extracted Document" directory in the base directory
        :param zip_dir_path: Path to the .zip file that is a copy of the .docx file
        :param doc_copy_path: Path to copy the original .docx file into the base directory
        """
        self.doc_path = doc_path
        self.base_dir_path = base_dir_path
        self.extracted_dir_path = extracted_dir_path
        self.zip_dir_path = zip_dir_path
        self.doc_copy_path = doc_copy_path

    def process_doc(self):
        """
        The process_doc function is responsible for:
            1. Copies original .docx file to base directory
            2. Extracts the documents contents by converting it to a .zip and
                unzipping it into "Extracted Document" directory
            3. Calls the itemize function in order to itemize the "Extracted Document" directory
        :return: None
        """
        # Copy the document to a new file that ends with .zip
        shutil.copy(self.doc_path, self.zip_dir_path)

        # Extract the .zip file into the "Extracted Document" directory
        with zipfile.ZipFile(self.zip_dir_path, "r") as zip_ref:
            zip_ref.extractall(self.extracted_dir_path)

        # Remove the unneeded zip file after it has been extracted
        os.remove(self.zip_dir_path)

        # Copy the original document into the base
        shutil.copy(self.doc_path, self.doc_copy_path)

        # Itemize the "Extracted Document" directory
        self.itemize()

    def itemize(self):
        """
        The itemize function is responsible for:
            1. Creates separate directories for these components of the document:
                    XML, CSS, Media, Content, RELS, and Uncatergorized
            2. Itemize the contents of the "Extracted Document" directory into the respective components directories
        :return: None
        """
        # Define the paths for the directories for each of the components
        xml_dir_path = os.path.join(self.base_dir_path, "XML")
        css_dir_path = os.path.join(self.base_dir_path, "CSS")
        media_dir_path = os.path.join(self.base_dir_path, "Media")
        content_dir_path = os.path.join(self.base_dir_path, "Content")
        uncategorized_dir_path = os.path.join(self.base_dir_path, "Uncategorized")
        rels_dir_path = os.path.join(self.base_dir_path, "RELS")

        # Create the directories for each the components
        os.mkdir(xml_dir_path)
        os.mkdir(css_dir_path)
        os.mkdir(media_dir_path)
        os.mkdir(content_dir_path)
        os.mkdir(uncategorized_dir_path)
        os.mkdir(rels_dir_path)

        # Loops through every directory and sub-directory that are in the "Extracted Document" directory
        for (dir_path, dir_names, file_names) in os.walk(self.extracted_dir_path):
            # Full path of the current directory
            current_dir_name = os.path.basename(dir_path)
            # For all files in the current directory
            for file_name in file_names:
                # Full path of the current file
                current_file_path = os.path.join(dir_path, file_name)
                # The "media" directory in the "Extracted Document" directory holds
                # certain content from the document such as images
                if current_dir_name == "media": # File is in the "media" directory
                    # Copy the current file into the media component's directory
                    media_file_path = os.path.join(media_dir_path, file_name)
                    shutil.copy(current_file_path, media_file_path)
                else:
                    # Get the document name without file extension and the file extension
                    doc_name, file_extension = os.path.splitext(file_name)
                    # Separate the files into their respective component's directory
                    if file_extension == ".xml":  # File is an XML file
                        # Path to copy the XML file to
                        xml_file_path = os.path.join(xml_dir_path, file_name)
                        # Copy the XML file into the XML component's directory
                        shutil.copy(current_file_path, xml_file_path)
                        # The "word" directory in the "Extracted Document" directory holds
                        # XML file that contain the user generated text
                        if current_dir_name == "word": # File is in the "word" directory
                            # Use XSLT to retrieve the user generated text from the XML
                            dom = et.parse(current_file_path)
                            # Use in-line XSLT instead of file so the script can run with one less dependency
                            xslt = et.fromstring(
                                "<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">\
                                <xsl:template match=\"*\"><xsl:value-of select=\".\"/></xsl:template></xsl:stylesheet>")
                            transform = et.XSLT(xslt)
                            new_dom = transform(dom)
                            # Remove the XML header from the text
                            output = str(new_dom).replace("<?xml version=\"1.0\"?>", "")
                            # Simple way of adding back newline characters to the text
                            # Adds a newline between a lowercase and uppercase letter. This does not work in all cases
                            for i in range(len(output) - 1):
                                current_char = output[i]
                                next_char = output[i + 1]
                                if (current_char.islower() and next_char.isupper()) \
                                        or (current_char.isdigit() and next_char.isupper()):
                                    output = output[:i + 1] + "\n" + output[i + 1:]
                            # Write the contents to a .txt file within the Content component's directory
                            word_file_path = os.path.join(content_dir_path, doc_name + ".txt")
                            text_file = open(word_file_path, "w")
                            text_file.write(output)
                            text_file.close()
                    elif file_extension == ".css":  # File is an CSS file
                        # Path to copy the CSS file to
                        css_file_path = os.path.join(css_dir_path, file_name)
                        # Copy the CSS file into the CSS component's directory
                        shutil.copy(current_file_path, css_file_path)
                    elif file_extension == ".rels":  # File is an RELS file
                        # Path to copy the RELS file to
                        rels_file_path = os.path.join(rels_dir_path, file_name)
                        # Copy the RELS file into the RELS component's directory
                        shutil.copy(current_file_path, rels_file_path)
                    else: # File is an uncategorized file
                        # Path to copy the uncategorized file to
                        uncategorized_file_path = os.path.join(uncategorized_dir_path, file_name)
                        # Copy the uncategorized file into the Uncategorized component's directory
                        shutil.copy(current_file_path, uncategorized_file_path)


class ImageFinder:

    def __init__(self, base_dir_path, extracted_dir_path):
        """
        The ImageFinder class is responsible for finding images that have the
        wrong extension that may be hidden in the document

        :param base_dir_path: Path to the base directory that holds all of the output of this script
        :param extracted_dir_path: Path to the "Extracted Document" directory in the base directory
        """
        self.hidden_images_dir_path = os.path.join(base_dir_path, "Hidden Images")
        self.extracted_dir_path = extracted_dir_path

    def get_hidden_images(self):
        """
        The get_hidden_images function is responsible for finding images that have the wrong extension that may be
        hidden in the document and returning the information about these files.

        Files are checked using the "imghdr" library to get the files' actual types

        :return: A list containing information about hidden image files
        (holds the original path, a path to a copy of the file in the "Images" directory,
         and a path to a copy of the file in the "Images" directory with the proper extension
        """
        # Create a list to hold all the information about the hidden images
        hidden_image_file_paths = []
        # Loops through every directory and sub-directory that are in the "Extracted Document" directory
        for (dir_path, dir_names, file_names) in os.walk(self.extracted_dir_path):
            # For all files in the current directory
            for file_name in file_names:
                current_file_path = os.path.join(dir_path, file_name)
                # Get the files current name and extension
                base_file_name, file_extension = os.path.splitext(file_name)
                # Get the file type according to imghdr. None is returned for non-images
                file_type = imghdr.what(current_file_path)
                # Check if the file is an image and is using the wrong extension
                if file_type is not None and not file_extension.endswith(file_type):
                    # Create a directory for the hidden images if it does not already exist
                    if not os.path.isdir(self.hidden_images_dir_path):
                        os.mkdir(self.hidden_images_dir_path)
                    # Copy the hidden image to the "Hidden Images" directory
                    hidden_image_file_path = os.path.join(self.hidden_images_dir_path, file_name)
                    shutil.copy(current_file_path, hidden_image_file_path)
                    # Copy the hidden image to the "Hidden Images" directory but add the proper extension
                    image_file_path = os.path.join(self.hidden_images_dir_path, base_file_name + "." + file_type)
                    shutil.copy(current_file_path, image_file_path)
                    # Store the hidden image's original path and the copies' paths into the hidden images list
                    hidden_image_file_paths.append([current_file_path, hidden_image_file_path, image_file_path])
        # Return a list contain all of the information about hidden images
        return hidden_image_file_paths


class Searcher:

    def __init__(self, base_dir_path, extracted_dir_path, search_term):
        """
        The Searcher class is responsible for finding all file's that
        contain the search term in their name or in their contents

        Creates a .txt file contain the search term in the "Search" directory in the base directory

        :param base_dir_path: Path to the base directory that holds all of the output of this script
        :param extracted_dir_path: Path to the "Extracted Document" directory in the base directory
        :param search_term: Path to the directory to put all files that are related to the search term
        """
        self.search_dir_path = os.path.join(base_dir_path, "Search")
        self.unreadable_dir_path = os.path.join(self.search_dir_path, "Unreadable")
        self.extracted_dir_path = extracted_dir_path
        self.search_term = search_term

        # Create a "Search" directory
        os.mkdir(self.search_dir_path)
        # Create a file named "search_term.txt" in the "Search" and write the search_term to it
        search_term_file = os.path.join(self.search_dir_path, "search_term.txt")
        file = open(search_term_file, "w")
        file.write(search_term)
        file.close()

    def find_search_term(self):

        """
        The find_search_term function is responsible for finding all file's that
        contain the search term in their name or in their contents

        Files are checked using the "re" library to get the files' actual types

        :return: A list containing information about files that
        contain the search term in their name or in their contents
        (holds the original paths,
        paths to a copies of the files in the "Search" directory,
        paths to files that could not be read,
        and paths to a copies of the unreadable file in the "Search/Unreadable" directory)
        """
        # Create lists to hold all the information for files related to the search feature
        found_file_names = []
        found_file_contents = []
        unreadable_file_contents = []
        search_file_paths = []
        unreadable_file_paths = []
        # Loops through every directory and sub-directory that are in the "Extracted Document" directory
        for (dir_path, dir_names, file_names) in os.walk(self.extracted_dir_path):
            # For all files in the current directory
            for file_name in file_names:
                # Path to the current file
                current_file_path = os.path.join(dir_path, file_name)
                # Path to copy the current file to into the "Search" directory in the base folder
                search_file_path = os.path.join(self.search_dir_path, file_name)
                # Use "re" library to perform regex match on the current file's name
                file_name_match = re.search(self.search_term, file_name)
                if file_name_match:  # Check if the file's name matches the regex
                    found_file_names.append(current_file_path)  # Add the file to the list of matching on file names
                    if not os.path.isfile(search_file_path):
                        # If the file has not been copied to the "Search" directory
                        # Copy the current file to the search directory
                        search_file_paths.append(search_file_path)
                        shutil.copy(current_file_path, search_file_path)
                try:
                    # Try to read contents of the current file
                    file = open(current_file_path, "r")
                    file_contents = file.read()
                    # Use "re" library to perform regex match on the current file's contents
                    file_contents_match = re.search(self.search_term, file_contents)
                    if file_contents_match:  # Check if the file's contents matches the regex
                        # Add the current file to the list of matching on file contents
                        found_file_contents.append(current_file_path)
                        if not os.path.isfile(search_file_path):
                            # If the current file has not been copied to the "Search" directory then
                            # copy the current file to the search directory
                            search_file_paths.append(search_file_path)
                            shutil.copy(current_file_path, search_file_path)
                    file.close()
                except UnicodeDecodeError:  # Except when the file is unreadable
                    if not os.path.isdir(self.unreadable_dir_path):
                        # Create "Unreadable" directory if it does not exist in the "Search" directory
                        os.mkdir(self.unreadable_dir_path)
                    if not os.path.isfile(search_file_path):
                        unreadable_file_path = os.path.join(self.unreadable_dir_path, file_name)
                        # Add the current file to the list of files with unreadable contents
                        unreadable_file_contents.append(current_file_path)
                        # Add the path to the copy of the current file to the list of copies of unreadable files
                        unreadable_file_paths.append(unreadable_file_path)
                        # Copy the current file to the "Unreadable" directory
                        shutil.copy(current_file_path, unreadable_file_path)
        # Return the needed information about files related to the search_term
        return found_file_names, found_file_contents, unreadable_file_contents, search_file_paths, unreadable_file_paths


def get_time_stamp():
    """
    The get_time_stamp function is a helper function used to get the current time stamp as a string
    Format: %Y%m%d-%H%M%S

    :return: The current time stamp as a string (Format: %Y%m%d-%H%M%S)
    """
    return time.strftime("%Y%m%d-%H%M%S")


def get_paths(doc_file_path):
    """
    The get_paths function is a helper function that is used to get the paths of directories
    that will be used throughout the script. All of these directories are based around the doc_file_path

    :param doc_file_path: Path to the original .docx file
    :return:    1. base_dir_path: Path to use for the base directory for the output files of this script
                2. extracted_dir_path: Path to use for the "Extracted Document" directory that will be in the base
                directory. This directory contains the unitemized files from unzipping the copy .docx file.
                3. zip_dir_path: Path to create the .zip file that is a copy of the .docx file
                4. doc_copy_path: Path to copy the original .docx file. This path is within the base directory.
                5. log_file_path: Path to create the log file
    """
    file_name = os.path.basename(doc_file_path)  # Get the document's file name from the path
    dir_path = os.path.dirname(doc_file_path)  # Get the path of the directory that contains the document
    doc_name = os.path.splitext(file_name)[0]  # Get the document's name without the extension

    time_stamp = get_time_stamp()  # Get the current time stamp as a string

    # Create the base directory name, "Extracted Directory" name, zip directory name, and log file name
    base_dir_name = doc_name + "_Itemized(" + time_stamp + ")"
    extracted_dir_name = "Extracted Document"
    zip_dir_name = base_dir_name + ".zip"
    log_file_name = "log.txt"

    # Create the base directory path, "Extracted Directory" path,
    # zip directory path, document copy path, and log file path
    base_dir_path = os.path.join(dir_path, base_dir_name)
    extracted_dir_path = os.path.join(base_dir_path, extracted_dir_name)
    zip_dir_path = os.path.join(dir_path, zip_dir_name)
    doc_copy_path = os.path.join(base_dir_path, file_name)
    log_file_path = os.path.join(base_dir_path, log_file_name)

    # Create the base directory
    os.mkdir(base_dir_path)

    # Return the paths
    return base_dir_path, extracted_dir_path, zip_dir_path, doc_copy_path, log_file_path


def log(log_file_path, message):
    """
    The log function is a helper function that is used to output the results to the console and a log file simultaneously

    :param log_file_path: Path to the log file
    :param message: The message to log
    :return: None
    """
    print(message)
    with open(log_file_path, "a") as file:
        file.write(message + "\n")


def run_docx_itemizer(doc_file_path, search_term, is_dir):
    """
    The run_docx_itemizer is responsible for running all of the sub classes and outputs each sub classes result.
        1. Runs the Itemizer class to extract and itemize the files from the document
        2. Runs the ImageFinder class to find any hidden images in the files from the document
        3. If a search_term is provided, runs the Searcher class to find any files that either their
        file name or contents regex match the search term

    :param doc_file_path: Path to the document
    :param search_term: The regex term to match with
    :param is_dir: True if the scripts is running on a whole directory. False if running on one file.
        This is used for better output formatting
    :return: None
    """

    # Get the  paths of directories that will be used throughout the script
    base_dir_path, extracted_dir_path, zip_dir_path, doc_copy_path, log_file_path = get_paths(doc_file_path)

    # Prefix is used to add an extra tab to the output if running on multiple documents. Imports ouputs formatting
    prefix = "\t" if is_dir else ""

    # Output general info about the file
    log(log_file_path, "Document Name: " + os.path.basename(doc_file_path))
    log(log_file_path, prefix + "Processing Document: " + os.path.abspath(doc_file_path))
    # Create an instance of the Itemizer and use it to process the document
    itemizer = Itemizer(doc_file_path, base_dir_path, extracted_dir_path, zip_dir_path, doc_copy_path)
    itemizer.process_doc()
    # The document has been completely itemized at this point
    # except for files related to hidden images, and search terms
    log(log_file_path, prefix + "Completed Itemizing Document")
    log(log_file_path, prefix + "Itemized Files Location: " + os.path.abspath(base_dir_path))
    log(log_file_path, prefix + "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    # Display info about an hidden images found in the document
    log(log_file_path, prefix + "Finding Hidden Images")
    # Create an instance of the ImageFinder and use it to find any hidden images in the document
    image_finder = ImageFinder(base_dir_path, extracted_dir_path)
    hidden_image_file_paths = image_finder.get_hidden_images()
    if len(hidden_image_file_paths) > 0: # Check if the document has any hidden images
        log(log_file_path, prefix + "Hidden Images Found:")
        for i in range(len(hidden_image_file_paths)): # For each hidden image log out it's info
            hidden_image_file_path = hidden_image_file_paths[i]
            log(log_file_path, prefix + "\tHidden Image File " + str(i + 1) + ": " + os.path.basename(hidden_image_file_path[0]))
            log(log_file_path, prefix + "\t\tHidden Image Found At: " + os.path.abspath(hidden_image_file_path[0]))
            log(log_file_path, prefix + "\t\tCopy Of Hidden Image At: " + os.path.abspath(hidden_image_file_path[1]))
            log(log_file_path, prefix + "\t\tCopy Of Hidden Image With Correct Extension: " + os.path.abspath(hidden_image_file_path[2]))
            if i < len(hidden_image_file_paths) - 1:
                log(log_file_path, prefix + "\t~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    else:
        log(log_file_path, prefix + "No Hidden Images Found")

    # Check to make sure there is a valid search term
    if search_term is not None and search_term is not "":
        # log out information about the search
        log(log_file_path, prefix + "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        log(log_file_path, prefix + "Searching File Names And Contents For: " + search_term)
        # Create an instance of the Searcher class and use it to search the documents files
        searcher = Searcher(base_dir_path, extracted_dir_path, search_term)
        found_file_names, found_file_contents, unreadable_file_contents, search_file_paths, unreadable_file_paths = searcher.find_search_term()
        if len(found_file_names) > 0: # List all files that their names match the regex search term
            log(log_file_path, prefix + "\tFile Names Match:")
            for file_path in found_file_names:
                log(log_file_path, prefix + "\t\t" + os.path.abspath(file_path))
        if len(found_file_contents) > 0: # List all files that their contents match the regex search term
            log(log_file_path, prefix + "\tFile Contents Match:")
            for file_path in found_file_contents:
                log(log_file_path, prefix + "\t\t" + os.path.abspath(file_path))
        if len(search_file_paths) > 0:
            # List all files that have been copied to the "Search" directory
            # because their names or contents matched the regex search term
            log(log_file_path, prefix + "\tFound Search Term Files Copied To:")
            for file_path in search_file_paths:
                log(log_file_path, prefix + "\t\t" + os.path.abspath(file_path))
        if len(unreadable_file_contents) > 0: # List all files that their contents could not be read
            log(log_file_path, prefix + "\tUnreadable File Contents:")
            for file_path in unreadable_file_contents:
                log(log_file_path, prefix + "\t\t" + os.path.abspath(file_path))
        if len(unreadable_file_paths) > 0:
            # List all files that have been copied to the "Search/Unreadable" directory
            # because their contents could not be read
            log(log_file_path, prefix + "\tUnreadable Files Copied To:")
            for file_path in unreadable_file_paths:
                log(log_file_path, prefix + "\t\t" + os.path.abspath(file_path))
        if len(found_file_names) == 0 and len(found_file_contents) == 0 and len(search_file_paths) == 0:
            log(log_file_path, prefix + "Search Term Not Found")
    log(log_file_path, "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")


def main():
    """
    The main function is responsible for handling the user provided arguments and calling the run_docx_itemizer
    accordingly. This included making sure the path is actually a file or directory, and that all files passed
    to the run_docx_itemizer function are .docx files.
    :return: None
    """
    # Get the user provided arguments from the argument parser
    args = parser.parse_args()
    path = args.path
    search_term = args.search_term

    # Check if the path is a directory
    if os.path.isdir(path):
        has_doc_file = False
        # For every directory and sub-directory in the directory at the "path"
        for (dir_path, dir_names, file_names) in os.walk(path):
            for file_name in file_names:  # For every file in the current directory
                file_extension = os.path.splitext(file_name)[1]
                if file_extension == ".docx":
                    # The file is a .docx file
                    current_file_path = os.path.join(dir_path, file_name)
                    # Run the docx itemizer
                    run_docx_itemizer(current_file_path, search_term, True)
                    has_doc_file = True
        if not has_doc_file:
            print("No .docx Files Found In: " + path)
    elif os.path.isfile(path): # The path is a single file
        file_extension = os.path.splitext(path)[1]
        if file_extension == ".docx": # The file is a .docx file
            # Run the docx itemizer
            run_docx_itemizer(path, search_term, False)
        else:
            print("File is not .docx: " + path)
    else:  # Path is not a directory or file
        print("File is not valid: " + path)


if __name__ == "__main__":
    main()
