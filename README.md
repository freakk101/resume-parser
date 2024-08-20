# Resume Parser with OpenAI

This Node.js application allows you to upload DOCX resume files and extract structured data such as employee names, technologies, and client names using OpenAI's GPT model. The extracted data is compiled into an Excel file for easy review.

## Features

- Upload multiple DOCX resumes at once.
- Extracts employee name, technology skills, and client names.
- Generates a combined Excel file with the extracted data.

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)
- [OpenAI API Key](https://platform.openai.com/api-keys)

## Setup Instructions (Windows)

1. Download and extract the app files from [GitHub](https://github.com/freakk101/resume-parser).
2. Open Command Prompt and navigate to the extracted folder.
3. Run `npm install` to install necessary dependencies.
4. Start the app using `node app.js`.
5. Open `http://localhost:3000` in your web browser.
6. Enter your OpenAI API key, upload your resumes, and process them.

## Usage

1. Enter your OpenAI API Key.
2. Upload DOCX resume files.
3. Click "Process Resumes" to generate the Excel file with structured resume data.

## Troubleshooting

- Ensure Node.js is installed.
- Make sure the app is running in the Command Prompt.
- Verify your API key is correct.

## License

This project is licensed under the MIT License.
