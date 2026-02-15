# MarkdownDoctor

Convert markdown text to DOCX format with an intuitive, easy-to-use web interface.

## Features

✨ **Simple & Intuitive** - Clean interface designed for ease of use  
📝 **Paste Markdown** - Directly paste markdown text (e.g., from ChatGPT)  
📤 **Upload Files** - Upload .md, .markdown, or .txt files  
👁️ **Live Preview** - See formatted preview before converting  
💾 **Instant Download** - Convert and download as .docx with one click

## Usage

1. **Open the tool**: Open `index.html` in your web browser
2. **Add your markdown**:
   - **Option A**: Paste markdown text directly into the text area
   - **Option B**: Upload a .md file using the "Upload File" tab
3. **Preview**: See a live preview of your formatted document
4. **Convert**: Click "Convert to DOCX" button
5. **Download**: Your .docx file will download automatically

## Running Locally

Simply open `index.html` in any modern web browser. No server required!

Or run with a local server:
```bash
python3 -m http.server 8080
# Then open http://localhost:8080 in your browser
```

## Technologies Used

- **markdown-it** - Markdown parser
- **docx** - DOCX document generation
- **file-saver** - File download functionality
- Pure HTML/CSS/JavaScript - No framework dependencies

## Perfect For

- Converting ChatGPT responses to Word documents
- Formatting markdown notes as DOCX files
- Quick document conversion without command-line tools
- Creating Word-compatible documents from markdown

## Browser Compatibility

Works in all modern browsers (Chrome, Firefox, Safari, Edge)
