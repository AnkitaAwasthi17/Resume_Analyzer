import PyPDF2
import subprocess
import os

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
        # Pipe the prompt as input to the ollama command
        result = subprocess.run(
            ["ollama", "run", "mistral"],
            input=prompt,          # Pass the prompt as input
            capture_output=True,   # Capture the output
            text=True,             # Ensure the output is returned as text
            encoding="utf-8"       # Use UTF-8 encoding
        )

        # Check if the command was successful
        if result.returncode == 0:
            return result.stdout.strip()  # Return the model's output (evaluation result)
        else:
            return f"Error: {result.stderr.strip()}"
    except FileNotFoundError:
        return "Error: The 'ollama' command was not found. Ensure it is installed and in the PATH."
    except Exception as e:
        return f"Exception: {str(e)}"

# Main function to process all PDFs in a folder and display results
def evaluate_all_cvs_in_folder(folder_path):
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    if not pdf_files:
        print(f"No PDF files found in folder '{folder_path}'.")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"\nProcessing file: {pdf_path}")

        cv_text = extract_text_from_pdf(pdf_path)
        if cv_text.startswith("Error reading PDF:"):
            print(cv_text)
            continue

        result = rate_cv_with_ollama(cv_text)
        if result.startswith("Error:") or result.startswith("Exception:"):
            print(f"Error processing {pdf_file}: {result}")
            continue

        # Display the result in the terminal
        print(f"\nModel response for {pdf_file}:\n{result}")
        
        
import csv

# Function to save the processed results into a CSV file
def save_results_to_csv(results, output_csv_path):
    # Define the column headers
    headers = [
        "Name", "Contact Details", "University", "Year of Study", "Course",
        "Discipline", "CGPA/Percentage", "Key Skills", "Gen AI Experience Score",
        "AI/ML Experience Score", "Supporting Information"
    ]

    # Open the CSV file for writing
    with open(output_csv_path, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=headers)

        # Write the header row
        writer.writeheader()

        # Write each row of results
        for result in results:
            writer.writerow(result)

# Function to process the model's response into a dictionary for CSV
def parse_model_response(response):
    data = {
        "Name": "",
        "Contact Details": "",
        "University": "",
        "Year of Study": "",
        "Course": "",
        "Discipline": "",
        "CGPA/Percentage": "",
        "Key Skills": "",
        "Gen AI Experience Score": "",
        "AI/ML Experience Score": "",
        "Supporting Information": ""
    }
    
    lines = response.splitlines()
    for line in lines:
        # Match known fields and map them to the corresponding dictionary keys
        if line.startswith("1. Name:"):
            data["Name"] = line.split(":", 1)[1].strip()
        elif line.startswith("2. Contact Details:"):
            data["Contact Details"] = line.split(":", 1)[1].strip()
        elif line.startswith("3. University:"):
            data["University"] = line.split(":", 1)[1].strip()
        elif line.startswith("4. Year of Study:"):
            data["Year of Study"] = line.split(":", 1)[1].strip()
        elif line.startswith("5. Course:"):
            data["Course"] = line.split(":", 1)[1].strip()
        elif line.startswith("6. Discipline:"):
            data["Discipline"] = line.split(":", 1)[1].strip()
        elif line.startswith("7. CGPA/Percentage:"):
            data["CGPA/Percentage"] = line.split(":", 1)[1].strip()
        elif line.startswith("8. Key Skills:"):
            data["Key Skills"] = line.split(":", 1)[1].strip()
        elif line.startswith("9. Gen AI Experience Score:"):
            data["Gen AI Experience Score"] = line.split(":", 1)[1].strip()
        elif line.startswith("10. AI/ML Experience Score:"):
            data["AI/ML Experience Score"] = line.split(":", 1)[1].strip()
        elif line.startswith("11. Supporting Information:"):
            data["Supporting Information"] += line.split(":", 1)[1].strip()
        elif line.startswith("-") or line.strip():  # Append extra info to Supporting Information
            data["Supporting Information"] += f" {line.strip()}"
    return data


# Main function to integrate with evaluation and save to CSV
def evaluate_and_save_to_csv(folder_path, output_csv_path):
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    if not pdf_files:
        print(f"No PDF files found in folder '{folder_path}'.")
        return

    results = []

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"\nProcessing file: {pdf_path}")

        cv_text = extract_text_from_pdf(pdf_path)
        if cv_text.startswith("Error reading PDF:"):
            print(cv_text)
            continue

        result = rate_cv_with_ollama(cv_text)
        if result.startswith("Error:") or result.startswith("Exception:"):
            print(f"Error processing {pdf_file}: {result}")
            continue

        # Parse the response into a dictionary and append to results
        parsed_data = parse_model_response(result)
        results.append(parsed_data)

    # Save all results to the CSV file
    save_results_to_csv(results, output_csv_path)
    print(f"\nResults saved to: {output_csv_path}")



if __name__ == "__main__":
    folder_path = "pdfs"  # Folder containing the PDFs
    output_csv_path = "processed_cvs.csv"  # Output CSV file path
    evaluate_and_save_to_csv(folder_path, output_csv_path)
