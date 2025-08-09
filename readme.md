PPT Inconsistency Analyzer


Project Description
This is a Python-based command-line tool that analyzes PowerPoint presentations (.pptx) to automatically detect factual and logical inconsistencies. The tool leverages the power of the Gemini API to compare content across all slides, identifying conflicting numerical data, contradictory claims, timeline mismatches, and other logical errors. This is particularly useful for fact-checking and quality assurance in presentations, reports, and pitch decks.

Features
Slide Content Extraction: Automatically extracts text and table data from each slide of a .pptx file.

AI-Powered Analysis: Utilizes the Gemini API to perform a detailed, cross-slide comparison of content.

Inconsistency Detection: Flags a variety of issues, including:

Conflicting numerical data (e.g., revenue figures, time savings).

Contradictory textual claims.

Logical mismatches in the overall narrative.

Structured Output: Provides a clear, well-formatted report in the terminal, referencing specific slide numbers and the nature of each issue.

2. The Solution: Explain the Approach
This project was designed and built in three main stages to handle a .pptx file from start to finish.

Breakdown of the Process :

Data Extraction: The python-pptx library is used to programmatically open the .pptx file. The script then iterates through each slide, extracting all available text and table data. This raw content is structured for easy analysis.

AI Analysis: The extracted content is sent to the Gemini API using a carefully crafted prompt. The prompt instructs the model to act as a specialized fact-checker, analyzing the data across all slides to identify inconsistencies and logical contradictions.

Report Generation: The script receives the structured response from the Gemini API and presents it as a clear, formatted report in the terminal. The report includes specific details about each inconsistency, including the slide numbers involved and the nature of the conflicting information.

Key Technical Choices :- 

Python: Chosen for its rich ecosystem of libraries, which simplifies complex tasks like file processing and API integration.

Libraries:

python-pptx: Handles the programmatic reading and parsing of .pptx files.

google-generativeai: Provides the interface for communicating with the Gemini API.

python-dotenv: A best practice library for securely managing environment variables like the API key, keeping them out of the codebase.

API Key Management: A dedicated .env file is used to store the Gemini API key, ensuring that sensitive information is not exposed in the source code.

Prerequisites
Before you begin, you need to have Python installed on your system. You will also need an API key for the Gemini API, which you can get for free from Google AI Studio.

Installation & Setup
1. Project Structure
Ensure your project directory is organized as follows:

ppt_analyzer/
├── presentations/
├── src/
│   └── main.py
├── .env
└── requirements.txt

2. Install Dependencies
Navigate to the ppt_analyzer/ directory in your terminal and install the required Python libraries using pip:

pip install python-pptx google-generativeai python-dotenv

Alternatively, you can list these in a requirements.txt file and run:

pip install -r requirements.txt

3. Configure the Gemini API Key
Create a file named .env in the root of your ppt_analyzer/ directory and add your API key in the following format:

GEMINI_API_KEY=your_gemini_api_key_here

Note: Replace your_gemini_api_key_here with your actual key. This keeps your key secure and separate from the code.



Limitations & Future Improvements
Image Analysis: The current version only processes text and tables. A major improvement would be to extract images from slides and include them in the Gemini API prompt to analyze inconsistencies within charts, diagrams, and other visual elements.

Large Decks: For very large presentations, the script could be optimized to handle API calls more efficiently, possibly by processing slides in chunks to avoid rate limits and improve performance.

Output Formats: The tool could be enhanced to output reports in different formats (e.g., JSON, HTML) or even automatically generate a new PowerPoint file with comments on the problematic slides.
