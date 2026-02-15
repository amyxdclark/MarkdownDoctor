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
const copyBtn = document.getElementById('copy-btn');
const clearBtn = document.getElementById('clear-btn');
const previewSection = document.getElementById('preview-section');
const previewContent = document.getElementById('preview-content');
const statusMessage = document.getElementById('status-message');

// Formatting options
const ignoreBoldCheckbox = document.getElementById('ignore-bold');
const ignoreItalicCheckbox = document.getElementById('ignore-italic');
const ignoreCodeCheckbox = document.getElementById('ignore-code');
const convertTablesCheckbox = document.getElementById('convert-tables');

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

// Copy button for email-ready output
copyBtn.addEventListener('click', async () => {
    if (!currentMarkdown.trim()) {
        showStatus('Please enter or upload markdown content', 'error');
        return;
    }
    
    try {
        // Apply formatting options and convert to plain text for email
        const formattedText = formatMarkdownForEmail(currentMarkdown);
        
        await navigator.clipboard.writeText(formattedText);
        showStatus('✓ Text copied to clipboard! Ready to paste into email.', 'success');
    } catch (error) {
        console.error('Copy error:', error);
        showStatus('Error copying to clipboard: ' + error.message, 'error');
    }
});

// Status message helpers
function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;
}

function hideStatus() {
    statusMessage.className = 'status-message';
}

// Format markdown for email (plain text with proper formatting)
function formatMarkdownForEmail(markdown) {
    let text = markdown;
    
    // Apply formatting options - process bold before italic to avoid conflicts
    if (ignoreBoldCheckbox.checked) {
        // Remove bold markers
        text = text.replace(/\*\*(.+?)\*\*/g, '$1');
        text = text.replace(/__(.+?)__/g, '$1');
    }
    
    if (ignoreItalicCheckbox.checked) {
        // Remove italic markers
        text = text.replace(/\*(.+?)\*/g, '$1');
        text = text.replace(/_(.+?)_/g, '$1');
    }
    
    if (ignoreCodeCheckbox.checked) {
        // Remove inline code markers
        text = text.replace(/`(.+?)`/g, '$1');
    }
    
    // Convert markdown tables to plain text if enabled
    if (convertTablesCheckbox.checked) {
        text = convertTablesToPlainText(text);
    }
    
    return text;
}

// Convert markdown tables to plain text representation
function convertTablesToPlainText(text) {
    const lines = text.split('\n');
    const result = [];
    let inTable = false;
    let tableRows = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Check if this is a table row
        if (line.includes('|')) {
            if (!inTable) {
                inTable = true;
                tableRows = [];
            }
            
            // Skip separator lines
            if (line.match(/^\|[\s\-:|]+\|$/)) {
                continue;
            }
            
            tableRows.push(line);
        } else {
            // End of table
            if (inTable && tableRows.length > 0) {
                result.push(formatTableAsPlainText(tableRows));
                tableRows = [];
                inTable = false;
            }
            result.push(line);
        }
    }
    
    // Handle table at end of file
    if (inTable && tableRows.length > 0) {
        result.push(formatTableAsPlainText(tableRows));
    }
    
    return result.join('\n');
}

// Format table rows as plain text
function formatTableAsPlainText(rows) {
    const parsedRows = rows.map(row => {
        return row.split('|')
            .map(cell => cell.trim())
            .filter(cell => cell.length > 0);
    });
    
    // Calculate column widths
    const colWidths = [];
    parsedRows.forEach(row => {
        row.forEach((cell, i) => {
            colWidths[i] = Math.max(colWidths[i] || 0, cell.length);
        });
    });
    
    // Format rows with padding
    const formattedRows = parsedRows.map((row, rowIndex) => {
        const formattedCells = row.map((cell, i) => {
            return cell.padEnd(colWidths[i], ' ');
        });
        return formattedCells.join(' | ');
    });
    
    // Add separator after header
    if (formattedRows.length > 0) {
        const separator = colWidths.map(w => '-'.repeat(w)).join('-+-');
        formattedRows.splice(1, 0, separator);
    }
    
    return formattedRows.join('\n');
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
    const { Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = docx;
    const lines = markdown.split('\n');
    const elements = [];
    let listItems = [];
    let inCodeBlock = false;
    let codeBlockContent = [];
    let inTable = false;
    let tableRows = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        // Handle tables if enabled
        if (convertTablesCheckbox.checked && line.includes('|')) {
            if (!inTable) {
                inTable = true;
                tableRows = [];
            }
            
            // Skip separator lines (must contain at least one dash or colon)
            if (line.match(/^\|?[\s]*[-:|]+[\s]*\|?$/) && line.includes('-')) {
                continue;
            }
            
            const cells = line.split('|')
                .map(cell => cell.trim())
                .filter(cell => cell.length > 0);
            
            if (cells.length > 0) {
                tableRows.push(cells);
            }
            continue;
        } else if (inTable && tableRows.length > 0) {
            // End of table, create table element
            elements.push(createTableElement(tableRows));
            tableRows = [];
            inTable = false;
        }
        
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
        } else if (line.startsWith('##### ')) {
            elements.push(new Paragraph({
                text: line.substring(6),
                heading: HeadingLevel.HEADING_5,
                spacing: { before: 150, after: 75 }
            }));
        } else if (line.startsWith('###### ')) {
            elements.push(new Paragraph({
                text: line.substring(7),
                heading: HeadingLevel.HEADING_6,
                spacing: { before: 140, after: 70 }
            }));
        } 
        // Handle blockquotes
        else if (line.match(/^>\s+/)) {
            const text = line.replace(/^>\s+/, '');
            const children = parseInlineFormatting(text);
            elements.push(new Paragraph({
                children: children,
                spacing: { before: 100, after: 100 },
                indent: { left: 720 },
                border: {
                    left: {
                        color: '10b981',
                        space: 1,
                        style: 'single',
                        size: 6
                    }
                }
            }));
        }
        // Handle horizontal rules
        else if (line.match(/^[\s]*(-{3,}|\*{3,}|_{3,})[\s]*$/)) {
            elements.push(new Paragraph({
                text: '',
                border: {
                    bottom: {
                        color: 'd1d5db',
                        space: 1,
                        style: 'single',
                        size: 6
                    }
                },
                spacing: { before: 200, after: 200 }
            }));
        }
        // Handle task lists
        else if (line.match(/^[\s]*-\s+\[(x| )\]\s+/i)) {
            const isChecked = line.match(/\[x\]/i);
            const text = line.replace(/^[\s]*-\s+\[(x| )\]\s+/i, '');
            const checkbox = isChecked ? '☑ ' : '☐ ';
            const children = parseInlineFormatting(text);
            elements.push(new Paragraph({
                children: [new TextRun({ text: checkbox }), ...children],
                spacing: { before: 100, after: 100 },
                indent: { left: 720 }
            }));
        }
        // Handle lists
        else if (line.match(/^[\s]*[-*]\s+/)) {
            const text = line.replace(/^[\s]*[-*]\s+/, '');
            const children = parseInlineFormatting(text);
            elements.push(new Paragraph({
                children: [new TextRun({ text: '• ' }), ...children],
                spacing: { before: 100, after: 100 },
                indent: { left: 720 }
            }));
        } else if (line.match(/^[\s]*\d+\.\s+/)) {
            const text = line.replace(/^[\s]*\d+\.\s+/, '');
            const num = line.match(/^[\s]*(\d+)\./)[1];
            const children = parseInlineFormatting(text);
            elements.push(new Paragraph({
                children: [new TextRun({ text: num + '. ' }), ...children],
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
    
    // Handle table at end of file
    if (inTable && tableRows.length > 0) {
        elements.push(createTableElement(tableRows));
    }
    
    return elements;
}

// Create a table element from rows
function createTableElement(rows) {
    const { Table, TableRow, TableCell, Paragraph, TextRun, WidthType, BorderStyle } = docx;
    
    const tableRows = rows.map((rowCells, rowIndex) => {
        const cells = rowCells.map(cellText => {
            return new TableCell({
                children: [new Paragraph({
                    children: parseInlineFormatting(cellText)
                })],
                width: {
                    size: Math.floor(9000 / rowCells.length),
                    type: WidthType.DXA
                }
            });
        });
        
        return new TableRow({
            children: cells
        });
    });
    
    return new Table({
        rows: tableRows,
        width: {
            size: 100,
            type: WidthType.PERCENTAGE
        }
    });
}

function parseInlineFormatting(text) {
    const { TextRun, ExternalHyperlink } = docx;
    const children = [];
    
    // First, handle links separately as they need special treatment
    // Extract and replace links with placeholders
    const linkPlaceholders = [];
    let textWithPlaceholders = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (match, linkText, url) => {
        const placeholder = `__LINK_${linkPlaceholders.length}__`;
        linkPlaceholders.push({ text: linkText, url: url });
        return placeholder;
    });
    
    // If all formatting is ignored, just return plain text
    if (ignoreBoldCheckbox.checked && ignoreItalicCheckbox.checked && ignoreCodeCheckbox.checked) {
        let plainText = textWithPlaceholders;
        // Process in order: bold, then italic, then code, then strikethrough
        plainText = plainText.replace(/\*\*(.+?)\*\*/g, '$1');
        plainText = plainText.replace(/__(.+?)__/g, '$1');
        plainText = plainText.replace(/\*(.+?)\*/g, '$1');
        plainText = plainText.replace(/_(.+?)_/g, '$1');
        plainText = plainText.replace(/`(.+?)`/g, '$1');
        plainText = plainText.replace(/~~(.+?)~~/g, '$1');
        
        // Restore links
        linkPlaceholders.forEach((link, i) => {
            plainText = plainText.replace(`__LINK_${i}__`, link.text);
        });
        
        children.push(new TextRun({ text: plainText }));
        return children;
    }
    
    // Parse inline formatting with enhanced regex - order matters: bold before italic
    let lastIndex = 0;
    
    // Combined regex to match bold, italic, strikethrough, code, and link placeholders
    const regex = /(\*\*|__|~~|`|\*|_|__LINK_\d+__)(.+?)(\1|__)/g;
    let match;
    
    while ((match = regex.exec(textWithPlaceholders)) !== null) {
        const marker = match[1];
        let content = match[2];
        const beforeText = textWithPlaceholders.substring(lastIndex, match.index);
        
        // Add text before the match
        if (beforeText) {
            children.push(new TextRun({ text: beforeText }));
        }
        
        // Handle link placeholders
        if (marker.startsWith('__LINK_')) {
            const linkIndex = parseInt(marker.match(/\d+/)[0]);
            const link = linkPlaceholders[linkIndex];
            // For DOCX, we'll just show the link text with underline to indicate it's a link
            children.push(new TextRun({ 
                text: link.text,
                underline: {},
                color: '0d9488'
            }));
            lastIndex = match.index + marker.length;
            continue;
        }
        
        // Add formatted text based on marker and options
        if (marker === '**' || marker === '__') {
            // Bold
            if (ignoreBoldCheckbox.checked) {
                children.push(new TextRun({ text: content }));
            } else {
                children.push(new TextRun({ text: content, bold: true }));
            }
        } else if (marker === '~~') {
            // Strikethrough
            children.push(new TextRun({ text: content, strike: true }));
        } else if (marker === '*' || marker === '_') {
            // Italic
            if (ignoreItalicCheckbox.checked) {
                children.push(new TextRun({ text: content }));
            } else {
                children.push(new TextRun({ text: content, italics: true }));
            }
        } else if (marker === '`') {
            // Code
            if (ignoreCodeCheckbox.checked) {
                children.push(new TextRun({ text: content }));
            } else {
                children.push(new TextRun({ 
                    text: content,
                    font: 'Courier New'
                }));
            }
        }
        
        lastIndex = regex.lastIndex;
    }
    
    // Add remaining text
    if (lastIndex < textWithPlaceholders.length) {
        let remaining = textWithPlaceholders.substring(lastIndex);
        
        // Check for any remaining link placeholders
        linkPlaceholders.forEach((link, i) => {
            remaining = remaining.replace(`__LINK_${i}__`, link.text);
        });
        
        children.push(new TextRun({ text: remaining }));
    }
    
    // If no matches found, return plain text (with links restored)
    if (children.length === 0) {
        let finalText = textWithPlaceholders;
        linkPlaceholders.forEach((link, i) => {
            finalText = finalText.replace(`__LINK_${i}__`, link.text);
        });
        children.push(new TextRun({ text: finalText }));
    }
    
    return children;
}
