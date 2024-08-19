const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { OpenAI } = require('openai');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

let openai;

// Function to extract text from DOCX
async function extractTextFromDocx(filePath) {
  const result = await mammoth.extractRawText({ path: filePath });
  return result.value;
}

// Function to structure resume data with OpenAI
async function structureResumeDataWithOpenAI(text, apiKey) {
  if (!openai) {
    openai = new OpenAI({ apiKey });
  }

  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        { role: "system", content: "You are a helpful assistant that extracts structured data from resumes. Always respond with valid JSON." },
        { role: "user", content: `Extract the following details from this resume and respond with ONLY the JSON, no other text:\n\n${text}\n\nDetails to extract:\n1. 'Employee Name': The name of the employee.\n2. 'Technology': A list of technologies, tools, and skills mentioned, presented as a single string.\n3. 'Client Names': A list of clients the employee has worked with, presented as a single string.` }
      ]
    });

    let contentString = response.choices[0].message.content.trim();
    
    // Remove any markdown formatting if present
    if (contentString.startsWith('```json')) {
      contentString = contentString.replace(/```json\n?/, '').replace(/\n?```$/, '');
    }

    // Attempt to parse the JSON
    try {
      return JSON.parse(contentString);
    } catch (parseError) {
      console.error('Error parsing OpenAI response:', parseError);
      throw new Error('Failed to parse OpenAI response as JSON');
    }
  } catch (error) {
    console.error('Error in OpenAI API call:', error);
    throw new Error('Failed to process resume with OpenAI: ' + error.message);
  }
}

// Enhanced React App
const App = () => {
  const [apiKey, setApiKey] = React.useState('');
  const [message, setMessage] = React.useState('');
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [fileCount, setFileCount] = React.useState(0);

  React.useEffect(() => {
    const savedApiKey = localStorage.getItem('openaiApiKey');
    if (savedApiKey) {
      setApiKey(savedApiKey);
    }
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    setError('');
    setMessage('');

    const formData = new FormData(e.target);
    localStorage.setItem('openaiApiKey', formData.get('apiKey'));

    try {
      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = 'combined_resumes.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setMessage('Processing complete. Excel file downloaded.');
      } else {
        const errorData = await response.json();
        throw new Error(`${errorData.error}${errorData.details ? ': ' + errorData.details : ''}`);
      }
    } catch (error) {
      setError(`Error: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileChange = (e) => {
    setFileCount(e.target.files.length);
  };

  return React.createElement(
    'div',
    { className: 'container' },
    React.createElement('h1', null, "Resume Processor with OpenAI"),
    React.createElement(
      'form',
      { onSubmit: handleSubmit },
      React.createElement(
        'div',
        { className: 'form-group' },
        React.createElement('label', { htmlFor: "apiKey" }, "OpenAI API Key:"),
        React.createElement('input', {
          type: "password",
          id: "apiKey",
          name: "apiKey",
          value: apiKey,
          onChange: (e) => setApiKey(e.target.value),
          required: true
        })
      ),
      React.createElement(
        'div',
        { className: 'form-group' },
        React.createElement('label', { htmlFor: "resumes" }, "Upload Resumes (DOCX files):"),
        React.createElement('input', {
          type: "file",
          id: "resumes",
          name: "resumes",
          multiple: true,
          accept: ".docx",
          required: true,
          onChange: handleFileChange
        })
      ),
      fileCount > 0 && React.createElement('p', null, `${fileCount} file(s) selected`),
      React.createElement(
        'button',
        { type: "submit", disabled: isLoading },
        isLoading ? "Processing..." : "Process Resumes"
      )
    ),
    isLoading && React.createElement('div', { className: 'loader' }),
    message && React.createElement('p', { className: 'message' }, message),
    error && React.createElement('p', { className: 'error' }, error)
  );
};

// Server-side rendering of the React app
const renderApp = () => {
  const app = ReactDOMServer.renderToString(React.createElement(App));
  return `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Resume Processor with OpenAI</title>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; padding: 20px; background-color: #f4f4f4; }
        .container { max-width: 600px; margin: 0 auto; background-color: #fff; padding: 20px; border-radius: 5px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1 { color: #333; text-align: center; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; color: #666; }
        input[type="password"], input[type="file"] { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; }
        button { background-color: #4CAF50; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; width: 100%; }
        button:disabled { background-color: #cccccc; }
        .loader { border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 20px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .message { color: green; text-align: center; }
        .error { color: red; text-align: center; }
      </style>
    </head>
    <body>
      <div id="root">${app}</div>
      <script src="https://unpkg.com/react@17/umd/react.development.js"></script>
      <script src="https://unpkg.com/react-dom@17/umd/react-dom.development.js"></script>
      <script>
        // Hydrate the app
        const App = ${App.toString()};
        ReactDOM.hydrate(
          React.createElement(App),
          document.getElementById('root')
        );
      </script>
    </body>
    </html>
  `;
};

// Routes
app.get('/', (req, res) => {
  res.send(renderApp());
});

app.post('/api/process', upload.array('resumes'), async (req, res) => {
  const apiKey = req.body.apiKey;
  const files = req.files;

  if (!apiKey) {
    return res.status(400).json({ error: 'API Key is required' });
  }

  if (!files || files.length === 0) {
    return res.status(400).json({ error: 'No files uploaded' });
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Resumes');

  worksheet.columns = [
    { header: 'File Name', key: 'fileName' },
    { header: 'Employee Name', key: 'name' },
    { header: 'Technology', key: 'technology' },
    { header: 'Client Name', key: 'client' },
    { header: 'Error', key: 'error' }
  ];

  let hasErrors = false;

  try {
    for (const file of files) {
      try {
        const text = await extractTextFromDocx(file.path);
        const structuredData = await structureResumeDataWithOpenAI(text, apiKey);

        worksheet.addRow({
          fileName: file.originalname,
          name: structuredData['Employee Name'] || 'Unknown',
          technology: structuredData['Technology'] || 'Unknown',
          client: structuredData['Client Names'] || 'Unknown',
          error: ''
        });
      } catch (fileError) {
        console.error(`Error processing file ${file.originalname}:`, fileError);
        worksheet.addRow({
          fileName: file.originalname,
          name: 'Error',
          technology: 'Error',
          client: 'Error',
          error: fileError.message
        });
        hasErrors = true;
      } finally {
        // Clean up uploaded file
        fs.unlinkSync(file.path);
      }
    }

    const excelFilePath = path.join(__dirname, 'combined_resumes.xlsx');
    await workbook.xlsx.writeFile(excelFilePath);

    res.download(excelFilePath, 'combined_resumes.xlsx', (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).json({ error: 'Error downloading file', details: err.message });
      }
      // Clean up Excel file after sending
      fs.unlinkSync(excelFilePath);
    });

    if (hasErrors) {
      console.log('Some files were processed with errors. Check the Excel file for details.');
    }
  } catch (error) {
    console.error('Error processing resumes:', error);
    res.status(500).json({ error: 'Error processing resumes', details: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});