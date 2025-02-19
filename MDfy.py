import mammoth
import os

# Specify the directory containing the .docx files
directory_path = r"C:\path\to\folder"  # Change this to your target directory

# Ensure the directory exists
if not os.path.exists(directory_path):
    print(f"Directory '{directory_path}' does not exist.")
    exit()

# Iterate over all files in the directory
for filename in os.listdir(directory_path):
    if filename.endswith(".docx"):
        docx_path = os.path.join(directory_path, filename)
        markdown_filename = f"{os.path.splitext(filename)[0]}.md"
        markdown_path = os.path.join(directory_path, markdown_filename)

        # Read and convert the DOCX file to Markdown
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_markdown(docx_file)

        # Write the converted content to a .md file
        with open(markdown_path, "w", encoding="utf-8") as markdown_file:
            markdown_file.write(result.value)

        print(f"Converted '{filename}' to '{markdown_filename}'")

print("Conversion completed.")