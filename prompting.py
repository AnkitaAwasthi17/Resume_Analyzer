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

    # Debug: Print the prompt being sent
    print("Prompt being sent to Mistral:\n", prompt)

    # Run the Ollama command using subprocess
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

# Function to save extracted data to an Excel file
def save_to_excel(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data)
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

# Main function to extract and process all PDFs in the folder
def evaluate_all_cvs_in_folder(folder_path):
    # Check if folder exists
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return

    # List all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]

    if not pdf_files:
        print(f"No PDF files found in folder '{folder_path}'.")
        return

    # Placeholder for all extracted data
    all_extracted_data = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Processing file: {pdf_path}")

        # Extract the text from the PDF
        cv_text = extract_text_from_pdf(pdf_path)

        if cv_text.startswith("Error reading PDF:"):
            print(cv_text)
            continue

        # Run the text through Ollama Mistral model to evaluate
        result = rate_cv_with_ollama(cv_text)

        if result.startswith("Error:") or result.startswith("Exception:"):
            print(f"Error processing {pdf_file}: {result}")
            continue

        # Parse the output from Mistral into structured data
        try:
            # Example response parsing - adapt based on actual response format
            extracted_data = {
                "Name": "Alex Johnson",
                "Contact Details": "alex.johnson@example.com, +112233445",
                "University": "DEF University",
                "Year of Study": "2nd Year",
                "Course": "B.Sc.",
                "Discipline": "Information Technology (IT)",
                "CGPA/Percentage": "7.8",
                "Key Skills": "Java, Deep Learning",
                "Gen AI Experience Score": "2 (Hands-on)",
                "AI/ML Experience Score": "1 (Exposed)",
                "Supporting Information": "Developed a generative art tool using GANs and participated in AI/ML hackathons",
            }
            # Add extracted data to the list
            all_extracted_data.append(extracted_data)
        except Exception as e:
            print(f"Error parsing response for {pdf_file}: {result}")
            continue

    # Save all extracted data to a single Excel file
    output_file = "evaluation_results_all.xlsx"
    save_to_excel(all_extracted_data, output_file)

# Example usage
if __name__ == "__main__":
    folder_path = "pdfs"  # Folder containing the PDFs
    evaluate_all_cvs_in_folder(folder_path)
