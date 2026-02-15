/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("convert").onclick = convertMarkdown;
  }
});

/**
 * Converts selected markdown text to Word format
 */
async function convertMarkdown() {
  try {
    await Word.run(async (context) => {
      // Get the current selection
      const selection = context.document.getSelection();
      
      // Load the text property
      selection.load("text");
      await context.sync();

      const markdownText = selection.text;
      
      if (!markdownText || markdownText.trim().length === 0) {
        showMessage("Please select some text first", "error");
        return;
      }

      // Parse and convert the markdown
      await convertMarkdownToWord(context, selection, markdownText);
      
      showMessage("Markdown converted successfully!", "success");
    });
  } catch (error) {
    console.error("Error converting markdown:", error);
    showMessage("Error: " + error.message, "error");
  }
}

/**
 * Parses markdown and applies Word formatting
 */
async function convertMarkdownToWord(context, selection, markdownText) {
  // Clear the selection first
  selection.clear();
  
  // Split into lines to process line by line
  const lines = markdownText.split('\n');
  let currentPosition = selection;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Skip empty lines
    if (line.trim().length === 0) {
      if (i < lines.length - 1) {
        currentPosition.insertParagraph("", Word.InsertLocation.end);
        currentPosition = context.document.getSelection();
      }
      continue;
    }
    
    // Process the line based on markdown syntax
    await processLine(context, currentPosition, line, i === 0);
    
    // Move to the next line position if not the last line
    if (i < lines.length - 1) {
      currentPosition = context.document.getSelection();
    }
  }
  
  await context.sync();
}

/**
 * Processes a single line of markdown
 */
async function processLine(context, position, line, isFirst) {
  // Headers (# H1, ## H2, etc.)
  const headerMatch = line.match(/^(#{1,6})\s+(.+)$/);
  if (headerMatch) {
    const level = headerMatch[1].length;
    const text = headerMatch[2];
    const paragraph = isFirst 
      ? position.insertText(text, Word.InsertLocation.replace)
      : position.insertParagraph(text, Word.InsertLocation.end);
    
    // Apply header style based on level
    paragraph.style = `Heading ${level}`;
    await context.sync();
    return;
  }
  
  // Unordered lists (- item or * item)
  const unorderedListMatch = line.match(/^[\s]*[-*]\s+(.+)$/);
  if (unorderedListMatch) {
    const text = unorderedListMatch[1];
    const processedText = processInlineMarkdown(text);
    const paragraph = isFirst
      ? position.insertText("", Word.InsertLocation.replace)
      : position.insertParagraph("", Word.InsertLocation.end);
    
    await applyInlineFormatting(context, paragraph, processedText);
    paragraph.leftIndent = 20;
    paragraph.firstLineIndent = -20;
    paragraph.spaceAfter = 0;
    await context.sync();
    return;
  }
  
  // Ordered lists (1. item, 2. item, etc.)
  const orderedListMatch = line.match(/^[\s]*\d+\.\s+(.+)$/);
  if (orderedListMatch) {
    const text = orderedListMatch[1];
    const processedText = processInlineMarkdown(text);
    const paragraph = isFirst
      ? position.insertText("", Word.InsertLocation.replace)
      : position.insertParagraph("", Word.InsertLocation.end);
    
    await applyInlineFormatting(context, paragraph, processedText);
    paragraph.leftIndent = 20;
    paragraph.spaceAfter = 0;
    await context.sync();
    return;
  }
  
  // Blockquotes (> quote)
  const blockquoteMatch = line.match(/^>\s+(.+)$/);
  if (blockquoteMatch) {
    const text = blockquoteMatch[1];
    const processedText = processInlineMarkdown(text);
    const paragraph = isFirst
      ? position.insertText("", Word.InsertLocation.replace)
      : position.insertParagraph("", Word.InsertLocation.end);
    
    await applyInlineFormatting(context, paragraph, processedText);
    paragraph.leftIndent = 30;
    paragraph.font.italic = true;
    paragraph.font.color = "#666666";
    await context.sync();
    return;
  }
  
  // Code blocks (```code```)
  if (line.trim().startsWith('```')) {
    // For simplicity, just show as monospace text
    const text = line.replace(/```/g, '');
    const paragraph = isFirst
      ? position.insertText(text, Word.InsertLocation.replace)
      : position.insertParagraph(text, Word.InsertLocation.end);
    
    paragraph.font.name = "Courier New";
    paragraph.font.color = "#333333";
    paragraph.shadingColor = "#f3f2f1";
    await context.sync();
    return;
  }
  
  // Regular paragraph with inline formatting
  const processedText = processInlineMarkdown(line);
  const paragraph = isFirst
    ? position.insertText("", Word.InsertLocation.replace)
    : position.insertParagraph("", Word.InsertLocation.end);
  
  await applyInlineFormatting(context, paragraph, processedText);
  await context.sync();
}

/**
 * Processes inline markdown (bold, italic, code, etc.)
 */
function processInlineMarkdown(text) {
  const segments = [];
  let currentPos = 0;
  
  // Regex patterns for inline markdown (order matters - more specific patterns first)
  const patterns = [
    { regex: /\*\*\*(.+?)\*\*\*/g, format: 'bold-italic' },
    { regex: /___(.+?)___/g, format: 'bold-italic' },
    { regex: /\*\*(.+?)\*\*/g, format: 'bold' },
    { regex: /__(.+?)__/g, format: 'bold' },
    { regex: /\*(.+?)\*/g, format: 'italic' },
    { regex: /_(.+?)_/g, format: 'italic' },
    { regex: /`(.+?)`/g, format: 'code' },
    { regex: /\[([^\]]+)\]\(([^)]+)\)/g, format: 'link' }
  ];
  
  // Find all matches with their positions
  const matches = [];
  patterns.forEach(pattern => {
    let match;
    // Reset regex lastIndex to start from beginning
    pattern.regex.lastIndex = 0;
    while ((match = pattern.regex.exec(text)) !== null) {
      matches.push({
        start: match.index,
        end: match.index + match[0].length,
        text: match[1],
        fullMatch: match[0],
        format: pattern.format,
        linkUrl: match[2] // For links
      });
    }
  });
  
  // Sort matches by position
  matches.sort((a, b) => a.start - b.start);
  
  // Remove overlapping matches (keep the first one)
  const filteredMatches = [];
  let lastEnd = -1;
  matches.forEach(match => {
    if (match.start >= lastEnd) {
      filteredMatches.push(match);
      lastEnd = match.end;
    }
  });
  
  // Build segments
  filteredMatches.forEach(match => {
    // Add text before the match
    if (match.start > currentPos) {
      segments.push({
        text: text.substring(currentPos, match.start),
        format: 'normal'
      });
    }
    
    // Add the formatted text
    segments.push({
      text: match.text,
      format: match.format,
      linkUrl: match.linkUrl
    });
    
    currentPos = match.end;
  });
  
  // Add remaining text
  if (currentPos < text.length) {
    segments.push({
      text: text.substring(currentPos),
      format: 'normal'
    });
  }
  
  return segments.length > 0 ? segments : [{ text: text, format: 'normal' }];
}

/**
 * Applies inline formatting to a paragraph
 */
async function applyInlineFormatting(context, paragraph, segments) {
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    const range = paragraph.insertText(segment.text, Word.InsertLocation.end);
    
    switch (segment.format) {
      case 'bold':
        range.font.bold = true;
        break;
      case 'italic':
        range.font.italic = true;
        break;
      case 'bold-italic':
        range.font.bold = true;
        range.font.italic = true;
        break;
      case 'code':
        range.font.name = "Courier New";
        range.font.color = "#d14";
        range.shadingColor = "#f3f2f1";
        break;
      case 'link':
        range.font.color = "#0078d4";
        range.font.underline = Word.UnderlineType.single;
        break;
    }
  }
  
  await context.sync();
}

/**
 * Shows a message to the user
 */
function showMessage(message, type) {
  const messageElement = document.getElementById("message");
  messageElement.textContent = message;
  messageElement.className = `message ${type}`;
  
  // Auto-hide success messages after 3 seconds
  if (type === "success") {
    setTimeout(() => {
      messageElement.className = "message";
    }, 3000);
  }
}
