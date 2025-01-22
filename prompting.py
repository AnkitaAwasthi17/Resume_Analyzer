import PyPDF2
import subprocess
import os
import pandas as pd

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    try:
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
            return text
    except Exception as e:
        return f"Error reading PDF: {str(e)}"

# Function to run Ollama Mistral locally via subprocess and get feedback
def rate_cv_with_ollama(cv_text):
    # Prepare the prompt for the Mistral model
    prompt = (
        f"Please analyze the following CV and extract the mandatory information in a structured format. "
        f"Provide scores for Generative AI Experience and AI/ML Experience based on the following scale: "
        f"1 – Exposed, 2 – Hands-on, 3 – Worked on advanced areas such as Agentic RAG, Evals, etc.\n\n"
        f"Mandatory Fields:\n"
        f"1. Name\n"
        f"2. Contact Details\n"
        f"3. University\n"
        f"4. Year of Study\n"
        f"5. Course\n"
        f"6. Discipline\n"
        f"7. CGPA/Percentage\n"
        f"8. Key Skills\n"
        f"9. Gen AI Experience Score\n"
        f"10. AI/ML Experience Score\n"
        f"11. Supporting Information (e.g., certifications, internships, projects)\n\n"
        f"Here is the CV text:\n\n{cv_text}"
    )

    try:
        result = subprocess.run(
            ["ollama", "run", "mistral"],
            input=prompt,
            capture_output=True,
            text=True,
            encoding="utf-8"
        )
        if result.returncode == 0:
            return result.stdout.strip()
        else:
            return f"Error: {result.stderr.strip()}"
    except FileNotFoundError:
        return "Error: The 'ollama' command was not found. Ensure it is installed and in the PATH."
    except Exception as e:
        return f"Exception: {str(e)}"

# Function to parse Mistral's response into structured data
def parse_response(response):
    try:
        # Split the response into key-value pairs
        lines = response.splitlines()
        data = {}
        for line in lines:
            if ":" in line:
                key, value = line.split(":", 1)
                data[key.strip()] = value.strip()
        return data
    except Exception as e:
        print(f"Error parsing response: {str(e)}")
        return None

# Function to save extracted data to an Excel file
def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

# Main function to extract and process all PDFs in the folder
def evaluate_all_cvs_in_folder(folder_path):
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    if not pdf_files:
        print(f"No PDF files found in folder '{folder_path}'.")
        return

    all_extracted_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Processing file: {pdf_path}")

        cv_text = extract_text_from_pdf(pdf_path)
        if cv_text.startswith("Error reading PDF:"):
            print(cv_text)
            continue

        result = rate_cv_with_ollama(cv_text)
        if result.startswith("Error:") or result.startswith("Exception:"):
            print(f"Error processing {pdf_file}: {result}")
            continue

        parsed_data = parse_response(result)
        if parsed_data:
            parsed_data["File Name"] = pdf_file  # Add file name for reference
            all_extracted_data.append(parsed_data)
        else:
            print(f"Failed to parse data for {pdf_file}")

    output_file = "evaluation_results_all.xlsx"
    save_to_excel(all_extracted_data, output_file)

# Example usage
if __name__ == "__main__":
    folder_path = "pdfs"  # Folder containing the PDFs
    evaluate_all_cvs_in_folder(folder_path)
