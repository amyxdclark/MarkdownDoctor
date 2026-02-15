// Initialize markdown parser
const md = window.markdownit();

// DOM Elements
const tabBtns = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');
const markdownInput = document.getElementById('markdown-input');
const uploadArea = document.getElementById('upload-area');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const removeFileBtn = document.getElementById('remove-file-btn');
const convertBtn = document.getElementById('convert-btn');
const clearBtn = document.getElementById('clear-btn');
const previewSection = document.getElementById('preview-section');
const previewContent = document.getElementById('preview-content');
const statusMessage = document.getElementById('status-message');

let currentMarkdown = '';
let currentFile = null;
const VALID_FILE_EXTENSIONS = ['.md', '.markdown', '.txt'];

// Tab switching
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        const tabName = btn.dataset.tab;
        
        // Update active tab button
        tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        // Update active tab content
        tabContents.forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`${tabName}-tab`).classList.add('active');
        
        // Clear status message
        hideStatus();
    });
});

// File upload handling
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

removeFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
});

function handleFile(file) {
    const fileName = file.name.toLowerCase();
    const isValid = VALID_FILE_EXTENSIONS.some(ext => fileName.endsWith(ext));
    
    if (!isValid) {
        showStatus('Please upload a valid markdown file (.md, .markdown, or .txt)', 'error');
        return;
    }
    
    currentFile = file;
    
    const reader = new FileReader();
    reader.onload = (e) => {
        currentMarkdown = e.target.result;
        
        // Show file info
        document.querySelector('.upload-content').style.display = 'none';
        fileInfo.style.display = 'flex';
        fileInfo.querySelector('.file-name').textContent = file.name;
        
        // Update preview
        updatePreview(currentMarkdown);
        hideStatus();
    };
    
    reader.onerror = () => {
        showStatus('Error reading file', 'error');
    };
    
    reader.readAsText(file);
}

function clearFile() {
    currentFile = null;
    currentMarkdown = '';
    fileInput.value = '';
    
    document.querySelector('.upload-content').style.display = 'block';
    fileInfo.style.display = 'none';
    
    previewSection.style.display = 'none';
    hideStatus();
}

// Markdown input handling
markdownInput.addEventListener('input', () => {
    currentMarkdown = markdownInput.value;
    if (currentMarkdown.trim()) {
        updatePreview(currentMarkdown);
    } else {
        previewSection.style.display = 'none';
    }
});

// Preview update
function updatePreview(markdown) {
    if (markdown.trim()) {
        const html = md.render(markdown);
        previewContent.innerHTML = html;
        previewSection.style.display = 'block';
    } else {
        previewSection.style.display = 'none';
    }
}

// Convert button
convertBtn.addEventListener('click', async () => {
    if (!currentMarkdown.trim()) {
        showStatus('Please enter or upload markdown content', 'error');
        return;
    }
    
    try {
        convertBtn.disabled = true;
        convertBtn.querySelector('.btn-text').textContent = 'Converting...';
        
        await convertMarkdownToDocx(currentMarkdown);
        
        showStatus('✓ Document converted successfully! Download started.', 'success');
    } catch (error) {
        console.error('Conversion error:', error);
        showStatus('Error converting document: ' + error.message, 'error');
    } finally {
        convertBtn.disabled = false;
        convertBtn.querySelector('.btn-text').textContent = 'Convert to DOCX';
    }
});

// Clear button
clearBtn.addEventListener('click', () => {
    markdownInput.value = '';
    currentMarkdown = '';
    clearFile();
    previewSection.style.display = 'none';
    hideStatus();
});

// Status message helpers
function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;
}

function hideStatus() {
    statusMessage.className = 'status-message';
}

// Markdown to DOCX conversion
async function convertMarkdownToDocx(markdown) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;
    
    // Parse markdown into sections
    const sections = parseMarkdown(markdown);
    
    // Create document
    const doc = new Document({
        sections: [{
            properties: {},
            children: sections
        }]
    });
    
    // Generate and download
    const blob = await Packer.toBlob(doc);
    const fileName = currentFile ? currentFile.name.replace(/\.(md|markdown|txt)$/i, '.docx') : 'converted-document.docx';
    saveAs(blob, fileName);
}

function parseMarkdown(markdown) {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    const lines = markdown.split('\n');
    const elements = [];
    let listItems = [];
    let inCodeBlock = false;
    let codeBlockContent = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        // Handle code blocks
        if (line.startsWith('```')) {
            if (inCodeBlock) {
                // End code block
                elements.push(new Paragraph({
                    children: [new TextRun({
                        text: codeBlockContent.join('\n'),
                        font: 'Courier New'
                    })],
                    spacing: { before: 200, after: 200 }
                }));
                codeBlockContent = [];
                inCodeBlock = false;
            } else {
                // Start code block
                inCodeBlock = true;
            }
            continue;
        }
        
        if (inCodeBlock) {
            codeBlockContent.push(line);
            continue;
        }
        
        // Handle headings
        if (line.startsWith('# ')) {
            elements.push(new Paragraph({
                text: line.substring(2),
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 240, after: 120 }
            }));
        } else if (line.startsWith('## ')) {
            elements.push(new Paragraph({
                text: line.substring(3),
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 100 }
            }));
        } else if (line.startsWith('### ')) {
            elements.push(new Paragraph({
                text: line.substring(4),
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 180, after: 90 }
            }));
        } else if (line.startsWith('#### ')) {
            elements.push(new Paragraph({
                text: line.substring(5),
                heading: HeadingLevel.HEADING_4,
                spacing: { before: 160, after: 80 }
            }));
        } 
        // Handle lists
        else if (line.match(/^[\s]*[-*]\s+/)) {
            const text = line.replace(/^[\s]*[-*]\s+/, '');
            elements.push(new Paragraph({
                text: '• ' + text,
                spacing: { before: 100, after: 100 },
                indent: { left: 720 }
            }));
        } else if (line.match(/^[\s]*\d+\.\s+/)) {
            const text = line.replace(/^[\s]*\d+\.\s+/, '');
            const num = line.match(/^[\s]*(\d+)\./)[1];
            elements.push(new Paragraph({
                text: num + '. ' + text,
                spacing: { before: 100, after: 100 },
                indent: { left: 720 }
            }));
        }
        // Handle regular paragraphs
        else if (line.trim() !== '') {
            const children = parseInlineFormatting(line);
            elements.push(new Paragraph({
                children: children,
                spacing: { before: 100, after: 100 }
            }));
        }
        // Handle empty lines
        else {
            elements.push(new Paragraph({
                text: '',
                spacing: { before: 120, after: 120 }
            }));
        }
    }
    
    return elements;
}

function parseInlineFormatting(text) {
    const { TextRun } = docx;
    const children = [];
    
    // TODO: Implement proper inline formatting (bold, italic, code)
    // Currently returns plain text with markdown syntax stripped
    let plainText = text;
    plainText = plainText.replace(/\*\*(.+?)\*\*/g, '$1');
    plainText = plainText.replace(/\*(.+?)\*/g, '$1');
    plainText = plainText.replace(/_(.+?)_/g, '$1');
    plainText = plainText.replace(/`(.+?)`/g, '$1');
    
    children.push(new TextRun({
        text: plainText
    }));
    
    return children;
}
