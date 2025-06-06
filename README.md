# New features in this fork:

- **Markdown conversion pipeline**: Full bidirectional conversion between Google Docs and Markdown
- **In-memory editing**: Edit documents as markdown strings with instruction-based modifications
- **Structure preservation**: Maintains heading hierarchy and paragraph formatting during conversion
- **Tab support**: Read from specific document tabs by name
- **Formatting preservation**: Advanced document editing with full text and paragraph styling support
- **Structure-aware editing**: Heading-based and section-aware tools for precise document editing

# Google Docs MCP Server

This is a Model Context Protocol (MCP) server that allows you to connect to Google Docs through Claude. With this server, you can:

- List all Google Docs in your Drive
- Read the content of specific documents (including support for tabbed documents)
- Read content from specific tabs by name
- List all tabs in a document
- Create new documents
- Update existing documents with advanced formatting
- Search for documents
- Delete documents
- **Structure-aware editing**: Find headings and edit content by document structure

## Markdown Conversion Pipeline

The server now includes a complete markdown conversion pipeline that implements the following workflow:

### 1. Document Reader Component
- **`read-doc-as-markdown`** - Converts Google Docs to structured Markdown
- Preserves heading hierarchy (H1-H6)
- Maintains paragraph structure and text content
- Handles tabbed documents

### 2. In-Memory Editing Component  
- **`edit-markdown-content`** - Edit Markdown strings using natural language instructions
- Supports instruction-based modifications like "add heading", "remove empty lines"
- Extensible framework for AI-powered content editing
- Preserves document structure during edits

### 3. Markdown Parser & Writer Component
- **`write-markdown-to-doc`** - Converts Markdown back to formatted Google Docs
- Automatically applies correct heading styles (HEADING_1, HEADING_2, etc.)
- Preserves paragraph formatting and text structure
- Clears existing content and replaces with formatted markdown

### Complete Workflow Example
```
1. read-doc-as-markdown(docId) → markdown string
2. edit-markdown-content(markdown, "add conclusion section") → edited markdown  
3. write-markdown-to-doc(docId, edited_markdown) → formatted Google Doc
```

This pipeline enables powerful document editing workflows where you can:
- Export documents to markdown for editing
- Apply AI-powered content modifications
- Reimport with full formatting preservation
- Work with documents programmatically using markdown syntax

## Structure-Aware Editing Tools

### Document Discovery
- **`find-headings`** - Discover all headings in a document with their exact positions
- Get structured view of document outline for precise editing

### Heading-Based Editing
- **`insert-content-after-heading`** - Insert content immediately after a specific heading
- **`replace-section-content`** - Replace all content in a section between headings
- **`append-to-section`** - Add content to the end of a specific section

These tools enable precise, structure-aware editing without losing document formatting. Perfect for:
- Adding content to specific sections
- Updating particular parts of long documents
- Maintaining document structure during edits
- Preserving existing formatting while making targeted changes

## Google Docs Tabs Support

This server fully supports the new [Google Docs tabs feature](https://developers.google.com/workspace/docs/api/how-tos/tabs). You can:
- List all tabs in a document
- Read content from a specific tab by name
- Work with nested child tabs
- Analyze individual tabs
- Apply structure-aware editing to specific tabs

## Text Formatting Support

The server provides comprehensive text formatting capabilities:

### Text Style Options
- **Bold, Italic, Underline, Strikethrough** - Basic text styling
- **Font Family** - Change to any supported font (Arial, Times New Roman, etc.)
- **Font Size** - Specify size in points
- **Text Color** - RGB color values (0-1 range)
- **Background Color** - Highlight text with background colors

### Paragraph Formatting
- **Alignment** - START, CENTER, END, JUSTIFIED
- **Line Spacing** - Single (1.0), 1.5x (1.5), Double (2.0), or custom
- **Space Above/Below** - Custom spacing in points

### Formatting Preservation
- **Auto-preserve** - Maintains existing formatting when inserting text
- **Style Reading** - Inspect current formatting at any location
- **Selective Styling** - Apply only specific formatting properties

## Prerequisites

- Node.js v16.0.0 or later
- Google Cloud project with the Google Docs API and Google Drive API enabled

## Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/google-docs-mcp-server.git
   cd google-docs-mcp-server
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Follow the [Google Cloud Console Setup](GOOGLE_SETUP_INSTRUCTIONS.md) to:**
   - Create a Google Cloud project
   - Enable the Google Docs and Drive APIs
   - Create OAuth 2.0 credentials
   - Download the `credentials.json` file

4. **Place your `credentials.json` file in the project root directory**

5. **Start the server:**
   ```bash
   npm start
   ```

The first time you run the server, it will open your browser for Google OAuth authentication.

## Available Tools

### Markdown Conversion Tools
- `read-doc-as-markdown` - Convert Google Doc to markdown format
- `edit-markdown-content` - Edit markdown content using natural language instructions  
- `write-markdown-to-doc` - Convert markdown back to formatted Google Doc

### Document Management
- `create-doc` - Create a new Google Doc
- `search-docs` - Search for documents by content
- `read-doc` - Read the full content of a document
- `delete-doc` - Delete a document
- `list-doc-tabs` - List all tabs in a document
- `read-doc-tab` - Read content from a specific tab

### Content Editing
- `update-doc` - Basic document updates (append or replace all)
- `update-doc-with-style` - Advanced updates with full formatting control
- `get-text-style` - Inspect formatting at specific locations

### Structure-Aware Editing
- `find-headings` - Discover document structure and heading positions
- `insert-content-after-heading` - Add content after specific headings
- `replace-section-content` - Replace content between headings  
- `append-to-section` - Add content to the end of sections

## Usage Examples

### Markdown Conversion Workflow
```
Convert document to markdown: read-doc-as-markdown with docId "abc123"
Edit the markdown: edit-markdown-content with instruction "add a conclusion section"
Write back to Google Docs: write-markdown-to-doc with the edited markdown
```

### Basic Document Operations
```
Create a new document titled "Meeting Notes"
Search for documents containing "quarterly report"  
Read the document with ID "abc123"
```

### Structure-Aware Editing
```
Find all headings in document ID "abc123"
Insert "New bullet points here" after the "Action Items" heading
Replace the content in the "Summary" section with updated text
Append additional notes to the "Conclusion" section
```

### Advanced Formatting
```
Update document 123 with "Executive Summary" using:
- Font: Arial, size 16
- Bold and underlined
- Blue text color
- Double line spacing
```

### Working with Tabs
```
List all tabs in document "abc123"
Read content from the "Appendix" tab
Insert content after "Results" heading in the "Data" tab
```

## Best Practices for Structure-Aware Editing

1. **Use `find-headings` first** to understand document structure
2. **Leverage section-based tools** instead of manual index calculations
3. **Preserve headings** when replacing section content
4. **Use consistent heading styles** for reliable structure detection
5. **Test with `read-doc` after edits** to verify results

The structure-aware tools make it safe to edit complex documents without breaking formatting or losing content. They work by:

- Analyzing document headings and structure
- Using semantic content boundaries instead of character positions
- Preserving existing formatting during edits
- Supporting both tabbed and regular documents

## Troubleshooting

If you encounter authentication issues:
1. Delete the `token.json` file in your project directory
2. Run the server again to trigger a new authentication flow

If you're having trouble with the Google Docs API:
1. Make sure the API is enabled in your Google Cloud Console
2. Check that your OAuth credentials have the correct scopes

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin feature/your-feature-name`
5. Submit a pull request

## License

MIT
