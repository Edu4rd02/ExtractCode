import os
from docx.shared import Pt
from doc_utils import add_uniform_heading

# Files and folders to ignore during extraction
IGNORE_LIST = {".git", "node_modules", "vendor", "storage", "resources",".gitattributes","assets","package-lock.json","composer.lock",".env","venv",".gitignore", ".env.local","cache"}

# File extensions to ignore during extraction
IGNORE_EXTENSIONS = {".pdf", ".log", ".ico", ".png"}

def is_ignored_file(name):
    return os.path.splitext(name)[1].lower() in IGNORE_EXTENSIONS

# Check if a directory contains at least one non-ignored file
def has_valid_content(path):
    try:
        entries = os.listdir(path)
    except OSError:
        return False
    for name in entries:
        if name in IGNORE_LIST:
            continue
        full = os.path.join(path, name)
        if os.path.isdir(full):
            if has_valid_content(full):
                return True
        elif not is_ignored_file(name):
            return True
    return False

# Recursive function to extract information from the specified path and add it to the document
def extractInfo(path, document, level=1, counters=None, base_parent=None):
    # Initialize counters for numbering the headings
    if counters is None:
        counters = [0] * 10
    # Set the base parent directory for calculating relative paths
    if base_parent is None:
        base_parent = os.path.dirname(path)

    # Iterate through the folders and files in the specified path
    for name in os.listdir(path):
        newPath = os.path.join(path, name)
        if name in IGNORE_LIST:
            continue
        elif is_ignored_file(name):
            continue

        is_dir = os.path.isdir(newPath)
        if is_dir:
            if not has_valid_content(newPath):
                continue

        # Increment the counter for the current level and reset counters for deeper levels
        counters[level - 1] += 1
        for i in range(level, len(counters)):
            counters[i] = 0

        # Create a numbered heading based on the current level and the counters
        number = ".".join(str(counters[i]) for i in range(level)) + "."
        rel_path = os.path.relpath(newPath, base_parent)
        heading_text = f"{number}  {rel_path}"

        # If the current path is a directory, add a heading and recursively extract information from it
        if is_dir:
            add_uniform_heading(document, heading_text, level)
            extractInfo(newPath, document, level + 1, counters, base_parent)
        # If the current path is a file, add a heading and read its content to add it to the document
        else:
            add_uniform_heading(document, heading_text, level)
            # Read the file content and add it to the document, ignoring any characters that are not valid for the document
            with open(newPath, "r", encoding="utf-8", errors="ignore") as file:
                content = file.read()
                content = "".join(ch for ch in content if ch >= " " or ch in "\n\r\t")
                # Add the file content in the document using Arial font with size 11
                code_para = document.add_paragraph('', style='Normal')
                code_run = code_para.add_run(content)
                code_run.font.name = 'Arial'
                code_run.font.size = Pt(11)