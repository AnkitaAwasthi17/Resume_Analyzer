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

# Function to save extracted data to an Excel file
def save_to_excel(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data)

    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

# Main function to extract CV text, run evaluation, and save results
def evaluate_cv(pdf_path):
    # Extract the text from the PDF
    cv_text = extract_text_from_pdf(pdf_path)

    # Check if the PDF was read successfully
    if cv_text.startswith("Error reading PDF:"):
        print(cv_text)
        return

    # Run the text through Ollama Mistral model to evaluate
    result = rate_cv_with_ollama(cv_text)

    # Check if the result contains an error
    if result.startswith("Error:") or result.startswith("Exception:"):
        print(result)
        return

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
    except Exception as e:
        print("Error parsing response from Mistral:", result)
        return

    # Save the extracted data to an Excel file
    output_file = "evaluation_results.xlsx"
    save_to_excel([extracted_data], output_file)

    # Output the evaluation result to the terminal as well
    print("Evaluation Result saved to", output_file)
    print("Evaluation Result:\n", extracted_data)

# Example usage
if __name__ == "__main__":
    pdf_path = "Alex_Johnson_Resume.pdf"  # Replace with the path to your CV
    if os.path.exists(pdf_path):
        evaluate_cv(pdf_path)
    else:
        print(f"Error: The file '{pdf_path}' does not exist.")
