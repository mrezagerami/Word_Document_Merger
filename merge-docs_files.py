#Merge Document Files
#Mohammad Reza (Arya) Gerami  - mr.gerami@gmail.com
import os
import re
import win32com.client

try:
    from docx import Document
except ImportError:
    print("Warning: 'python-docx' library is not installed. .docx files might not be read correctly.")

def natural_sort_key(filename):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', filename)]

def combine_all_word_files(folder_path, output_filename):
    try:
        all_files = os.listdir(folder_path)
    except FileNotFoundError:
        print(f"Error: The folder '{folder_path}' was not found.")
        return

    valid_extensions = ('.doc', '.docx')
    target_files = [f for f in all_files if f.lower().endswith(valid_extensions) and not f.startswith('~')]

    if not target_files:
        print("No .doc or .docx files found in the directory.")
        return

    target_files.sort(key=natural_sort_key)

    print(f"Found {len(target_files)} files. Files will be processed in this correct order:")
    for f in target_files:
        print(f" - {f}")
    print("-" * 40)

    word_app = None
    if any(f.lower().endswith('.doc') for f in target_files):
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
        except Exception as e:
            print(f"Error launching Word: {e}")

    abs_folder_path = os.path.abspath(folder_path)

    with open(output_filename, 'w', encoding='utf-8') as outfile:
        for filename in target_files:
            file_path = os.path.join(abs_folder_path, filename)
            
            outfile.write(f"\n\n{'='*40}\n")
            outfile.write(f"--- Document: {filename} ---\n")
            outfile.write(f"{'='*40}\n\n")

            if filename.lower().endswith('.docx'):
                try:
                    doc = Document(file_path)
                    for paragraph in doc.paragraphs:
                        outfile.write(paragraph.text + '\n')
                    print(f"Processed (.docx): {filename}")
                except Exception as e:
                    print(f"Failed to read {filename}: {e}")

            # پردازش فایل‌های .doc (قدیمی)
            elif filename.lower().endswith('.doc'):
                if word_app:
                    try:
                        doc = word_app.Documents.Open(file_path)
                        text = doc.Content.Text
                        outfile.write(text)
                        doc.Close(False)
                        print(f"Processed (.doc) : {filename}")
                    except Exception as e:
                        print(f"Failed to read {filename}: {e}")
                else:
                    print(f"Skipped {filename} - MS Word is not running.")

    if word_app:
        word_app.Quit()
        
    print(f"\nSuccess! All files have been correctly combined into '{output_filename}'.")

# ==========================================
# Example Usage
# ==========================================
if __name__ == "__main__":
    FOLDER_TO_SCAN = 'D:/ExampelFolder'  
    FINAL_OUTPUT_FILE = 'D:/ExampelFolder/combined_document.txt' 
    
    combine_all_word_files(FOLDER_TO_SCAN, FINAL_OUTPUT_FILE)