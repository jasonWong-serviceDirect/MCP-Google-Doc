import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { google } from "googleapis";
import { authenticate } from "@google-cloud/local-auth";
import * as fs from "fs";
import * as path from "path";
import * as process from "process";
import { z } from "zod";
import { docs_v1, drive_v3 } from "googleapis";
import { OAuth2Client } from "google-auth-library";

// Set up OAuth2.0 scopes - we need full access to Docs and Drive
const SCOPES = [
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.readonly" // Add read-only scope as a fallback
];

// Resolve paths relative to the project root
// Fix path resolution for Windows by removing the leading slash from file URLs
const currentFilePath = new URL(import.meta.url).pathname;
const fixedPath = process.platform === 'win32' && currentFilePath.startsWith('/') 
  ? currentFilePath.slice(1) 
  : currentFilePath;
const PROJECT_ROOT = path.resolve(path.join(path.dirname(fixedPath), '..'));

// The token path is where we'll store the OAuth credentials
const TOKEN_PATH = path.join(PROJECT_ROOT, "token.json");

// The credentials path is where your OAuth client credentials are stored
const CREDENTIALS_PATH = path.join(PROJECT_ROOT, "credentials.json");

// Create an MCP server instance
const server = new McpServer({
  name: "google-docs",
  version: "1.0.0",
});

/**
 * Load saved credentials if they exist, otherwise trigger the OAuth flow
 */
async function authorize() {
  try {
    // Load client secrets from a local file
    console.error("Reading credentials from:", CREDENTIALS_PATH);
    const content = fs.readFileSync(CREDENTIALS_PATH, "utf-8");
    const keys = JSON.parse(content);
    const clientId = keys.installed.client_id;
    const clientSecret = keys.installed.client_secret;
    const redirectUri = keys.installed.redirect_uris[0];
    
    console.error("Using client ID:", clientId);
    console.error("Using redirect URI:", redirectUri);
    
    // Create an OAuth2 client
    const oAuth2Client = new OAuth2Client(clientId, clientSecret, redirectUri);
    
    // Check if we have previously stored a token
    if (fs.existsSync(TOKEN_PATH)) {
      console.error("Found existing token, attempting to use it...");
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf-8"));
      oAuth2Client.setCredentials(token);
      return oAuth2Client;
    }
    
    // No token found, use the local-auth library to get one
    console.error("No token found, starting OAuth flow...");
    const client = await authenticate({
      scopes: SCOPES,
      keyfilePath: CREDENTIALS_PATH,
    });
    
    if (client.credentials) {
      console.error("Authentication successful, saving token...");
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(client.credentials));
      console.error("Token saved successfully to:", TOKEN_PATH);
    } else {
      console.error("Authentication succeeded but no credentials returned");
    }
    
    return client;
  } catch (err) {
    console.error("Error authorizing with Google:", err);
    if (err.message) console.error("Error message:", err.message);
    if (err.stack) console.error("Stack trace:", err.stack);
    throw err;
  }
}

// Create Docs and Drive API clients
let docsClient: docs_v1.Docs;
let driveClient: drive_v3.Drive;

// Initialize Google API clients
async function initClients() {
  try {
    console.error("Starting client initialization...");
    const auth = await authorize();
    console.error("Auth completed successfully:", !!auth);
    docsClient = google.docs({ version: "v1", auth: auth as any });
    console.error("Docs client created:", !!docsClient);
    driveClient = google.drive({ version: "v3", auth: auth as any });
    console.error("Drive client created:", !!driveClient);
    return true;
  } catch (error) {
    console.error("Failed to initialize Google API clients:", error);
    return false;
  }
}

// Initialize clients when the server starts
initClients().then((success) => {
  if (!success) {
    console.error("Failed to initialize Google API clients. Server will not work correctly.");
  } else {
    console.error("Google API clients initialized successfully.");
  }
});

// RESOURCES

// Resource for listing documents
server.resource(
  "list-docs",
  "googledocs://list",
  async (uri) => {
    try {
      const response = await driveClient.files.list({
        q: "mimeType='application/vnd.google-apps.document'",
        fields: "files(id, name, createdTime, modifiedTime)",
        pageSize: 50,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });

      const files = response.data.files || [];
      let content = "Google Docs in your Drive:\n\n";
      
      if (files.length === 0) {
        content += "No Google Docs found.";
      } else {
        files.forEach((file: any) => {
          content += `Title: ${file.name}\n`;
          content += `ID: ${file.id}\n`;
          content += `Created: ${file.createdTime}\n`;
          content += `Last Modified: ${file.modifiedTime}\n\n`;
        });
      }

      return {
        contents: [{
          uri: uri.href,
          text: content,
        }]
      };
    } catch (error) {
      console.error("Error listing documents:", error);
      return {
        contents: [{
          uri: uri.href,
          text: `Error listing documents: ${error}`,
        }]
      };
    }
  }
);

// Resource to get a specific document by ID
server.resource(
  "get-doc",
  new ResourceTemplate("googledocs://{docId}", { list: undefined }),
  async (uri, { docId }) => {
    try {
      const doc = await docsClient.documents.get({
        documentId: docId as string,
      });
      
      // Extract the document content
      let content = `Document: ${doc.data.title}\n\n`;
      
      // Process the document content from the complex data structure
      const document = doc.data;
      if (document && document.body && document.body.content) {
        let textContent = "";
        
        // Loop through the document's structural elements
        document.body.content.forEach((element: any) => {
          if (element.paragraph) {
            element.paragraph.elements.forEach((paragraphElement: any) => {
              if (paragraphElement.textRun && paragraphElement.textRun.content) {
                textContent += paragraphElement.textRun.content;
              }
            });
          }
        });
        
        content += textContent;
      }

      return {
        contents: [{
          uri: uri.href,
          text: content,
        }]
      };
    } catch (error) {
      console.error(`Error getting document ${docId}:`, error);
      return {
        contents: [{
          uri: uri.href,
          text: `Error getting document ${docId}: ${error}`,
        }]
      };
    }
  }
);

// TOOLS

// Tool to create a new document
server.tool(
  "create-doc",
  {
    title: z.string().describe("The title of the new document"),
    content: z.string().optional().describe("Optional initial content for the document"),
  },
  async ({ title, content = "" }) => {
    try {
      // Create a new document
      const doc = await docsClient.documents.create({
        requestBody: {
          title: title,
        },
      });

      const documentId = doc.data.documentId;

      // If content was provided, add it to the document
      if (content) {
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: {
                    index: 1,
                  },
                  text: content,
                },
              },
            ],
          },
        });
      }

      return {
        content: [
          {
            type: "text",
            text: `Document created successfully!\nTitle: ${title}\nDocument ID: ${documentId}\nYou can now reference this document using: googledocs://${documentId}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error creating document:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);





// Tool to search for documents
server.tool(
  "search-docs",
  {
    query: z.string().describe("The search query to find documents"),
  },
  async ({ query }) => {
    try {
      const response = await driveClient.files.list({
        q: `mimeType='application/vnd.google-apps.document' and fullText contains '${query}'`,
        fields: "files(id, name, createdTime, modifiedTime)",
        pageSize: 10,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });
      
      // Add response logging for debugging
      console.error("Drive API Response:", JSON.stringify(response, null, 2));
      
      // Add better response validation
      if (!response || !response.data) {
        throw new Error("Invalid response from Google Drive API");
      }
      
      // Add null check and default to empty array
      const files = (response.data.files || []);
      
      let content = `Search results for "${query}":\n\n`;
      
      if (files.length === 0) {
        content += "No documents found matching your query.";
      } else {
        files.forEach((file: any) => {
          content += `Title: ${file.name}\n`;
          content += `ID: ${file.id}\n`;
          content += `Created: ${file.createdTime}\n`;
          content += `Last Modified: ${file.modifiedTime}\n\n`;
        });
      }
      
      return {
        content: [
          {
            type: "text",
            text: content,
          },
        ],
      };
    } catch (error) {
      console.error("Error searching documents:", error);
      // Include more detailed error information
      const errorMessage = error instanceof Error 
          ? `${error.message}\n${error.stack}` 
          : String(error);
          
      return {
        content: [
          {
            type: "text",
            text: `Error searching documents: ${errorMessage}`,
          },
        ],
        isError: true,
      };
    }
  }
);



// Helper function to extract text content from document elements
function extractTextFromContent(content: any[]): string {
  let textContent = "";
  
  content.forEach((element: any) => {
    if (element.paragraph) {
      element.paragraph.elements.forEach((paragraphElement: any) => {
        if (paragraphElement.textRun && paragraphElement.textRun.content) {
          textContent += paragraphElement.textRun.content;
        }
      });
    } else if (element.table) {
      // Handle table content
      element.table.tableRows.forEach((row: any) => {
        row.tableCells.forEach((cell: any) => {
          if (cell.content) {
            textContent += extractTextFromContent(cell.content);
          }
        });
      });
    }
  });
  
  return textContent;
}

// Helper function to collect all tabs (including child tabs) recursively
function collectAllTabs(tabs: any[]): any[] {
  const allTabs: any[] = [];
  
  function collectTabs(tabList: any[]) {
    tabList.forEach((tab: any) => {
      if (tab.documentTab) {
        allTabs.push(tab);
      }
      if (tab.childTabs && tab.childTabs.length > 0) {
        collectTabs(tab.childTabs);
      }
    });
  }
  
  collectTabs(tabs);
  return allTabs;
}



// Tool to list all tabs in a document
server.tool(
  "list-doc-tabs",
  {
    docId: z.string().describe("The ID of the document to list tabs for"),
  },
  async ({ docId }) => {
    try {
      // Use the googleapis client but bypass TypeScript checking for the new parameter
      // The includeTabsContent parameter is officially supported by the Google Docs API
      // Reference: https://developers.google.com/workspace/docs/api/how-tos/tabs
      const response = await (docsClient.documents as any).get({
        documentId: docId,
        includeTabsContent: true,
      });
      
      const doc = response.data;
      let content = `Document: ${doc.title}\n\n`;
      
      // Check if document has tabs
      if (!doc.tabs || doc.tabs.length === 0) {
        content += "This document does not have any tabs.";
        return {
          content: [
            {
              type: "text",
              text: content,
            },
          ],
        };
      }
      
      // Collect all tabs (including child tabs)
      const allTabs = collectAllTabs(doc.tabs);
      
      content += `Found ${allTabs.length} tab(s):\n\n`;
      
      allTabs.forEach((tab: any, index: number) => {
        const tabTitle = tab.tabProperties?.title || "Untitled Tab";
        const tabId = tab.tabProperties?.tabId || "No ID";
        content += `${index + 1}. "${tabTitle}" (ID: ${tabId})\n`;
      });

      return {
        content: [
          {
            type: "text",
            text: content,
          },
        ],
      };
    } catch (error) {
      console.error(`Error listing tabs for document ${docId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error listing tabs for document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to delete a document
server.tool(
  "delete-doc",
  {
    docId: z.string().describe("The ID of the document to delete"),
  },
  async ({ docId }) => {
    try {
      // Get the document title first for confirmation
      const doc = await docsClient.documents.get({ documentId: docId });
      const title = doc.data.title;
      
      // Delete the document
      await driveClient.files.delete({
        fileId: docId,
      });

      return {
        content: [
          {
            type: "text",
            text: `Document "${title}" (ID: ${docId}) has been successfully deleted.`,
          },
        ],
      };
    } catch (error) {
      console.error(`Error deleting document ${docId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error deleting document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);



// =============================================
// MARKDOWN PARSING AND CONVERSION TOOLS
// =============================================





// PROMPTS

// Prompt for document creation
server.prompt(
  "create-doc-template",
  { 
    title: z.string().describe("The title for the new document"),
    subject: z.string().describe("The subject/topic the document should be about"),
    style: z.string().describe("The writing style (e.g., formal, casual, academic)"),
  },
  ({ title, subject, style }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please create a Google Doc with the title "${title}" about ${subject} in a ${style} writing style. Make sure it's well-structured with an introduction, main sections, and a conclusion.`
      }
    }]
  })
);

// Prompt for document analysis
server.prompt(
  "analyze-doc",
  { 
    docId: z.string().describe("The ID of the document to analyze"),
  },
  ({ docId }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please analyze the content of the document with ID ${docId}. Provide a summary of its content, structure, key points, and any suggestions for improvement.`
      }
    }]
  })
);

// Prompt for analyzing a specific tab
server.prompt(
  "analyze-doc-tab",
  { 
    docId: z.string().describe("The ID of the document to analyze"),
    tabName: z.string().describe("The name of the tab to analyze"),
  },
  ({ docId, tabName }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please analyze the content of the tab "${tabName}" in the document with ID ${docId}. Provide a summary of the tab's content, key points, and any suggestions for improvement specific to this tab.`
      }
    }]
  })
);

// =============================================
// CORE MARKDOWN CONVERSION COMPONENTS
// =============================================

/**
 * Document Reader Component
 * Converts Google Docs to Markdown with structure preserved
 */
function convertDocToMarkdown(docContent: any): string {
  if (!docContent || !docContent.body || !docContent.body.content) {
    return "";
  }

  let markdown = "";
  
  function traverseDocContent(content: any[]): string {
    let result = "";
    
    for (const element of content) {
      if (element.paragraph) {
        const paragraph = element.paragraph;
        let paragraphText = "";
        
        // Extract text from paragraph elements
        if (paragraph.elements) {
          for (const paragraphElement of paragraph.elements) {
            if (paragraphElement.textRun && paragraphElement.textRun.content) {
              paragraphText += paragraphElement.textRun.content;
            }
          }
        }
        
        // Convert paragraph style to markdown
        const namedStyleType = paragraph.paragraphStyle?.namedStyleType;
        
        switch (namedStyleType) {
          case 'HEADING_1':
            result += `# ${paragraphText.trim()}\n\n`;
            break;
          case 'HEADING_2':
            result += `## ${paragraphText.trim()}\n\n`;
            break;
          case 'HEADING_3':
            result += `### ${paragraphText.trim()}\n\n`;
            break;
          case 'HEADING_4':
            result += `#### ${paragraphText.trim()}\n\n`;
            break;
          case 'HEADING_5':
            result += `##### ${paragraphText.trim()}\n\n`;
            break;
          case 'HEADING_6':
            result += `###### ${paragraphText.trim()}\n\n`;
            break;
          default:
            // NORMAL_TEXT or other styles
            if (paragraphText.trim()) {
              result += `${paragraphText.trim()}\n\n`;
            }
            break;
        }
      } else if (element.table) {
        // Handle tables (basic implementation)
        result += "<!-- Table content -->\n\n";
      } else if (element.tableOfContents) {
        // Handle table of contents
        result += "<!-- Table of Contents -->\n\n";
      }
    }
    
    return result;
  }
  
  markdown = traverseDocContent(docContent.body.content);
  
  // Clean up extra newlines
  return markdown.replace(/\n{3,}/g, '\n\n').trim();
}



/**
 * Enhanced Markdown Parser Component
 * Takes Markdown and produces a list of styled segments for Google Docs
 */
function parseMarkdownToDocRequests(markdownString: string, startIndex: number = 1, tabId?: string): any[] {
  const lines = markdownString.split('\n');
  const requests: any[] = [];
  let currentIndex = startIndex;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Skip empty lines but preserve them in output
    if (!line.trim()) {
      if (i < lines.length - 1) { // Don't add newline for last empty line
        requests.push({
          insertText: {
            location: { 
              index: currentIndex,
              ...(tabId && { tabId })
            },
            text: '\n'
          }
        });
        currentIndex += 1;
      }
      continue;
    }
    
    let namedStyleType = 'NORMAL_TEXT';
    let text = line;
    
    // Parse markdown headings
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      text = headingMatch[2].trim();
      namedStyleType = `HEADING_${level}`;
    }
    
    // Parse inline formatting and create styled segments
    const segments = parseInlineFormatting(text);
    
    // Process each segment
    for (const segment of segments) {
      const textWithNewline = (segment === segments[segments.length - 1]) ? segment.text + '\n' : segment.text;
      
      // Insert text request
      requests.push({
        insertText: {
          location: { 
            index: currentIndex,
            ...(tabId && { tabId })
          },
          text: textWithNewline
        }
      });
      
      const segmentEndIndex = currentIndex + textWithNewline.length;
      
      // Apply paragraph style to the entire line (only for the first segment)
      if (segment === segments[0]) {
        requests.push({
          updateParagraphStyle: {
            range: {
              startIndex: currentIndex,
              endIndex: currentIndex + segments.reduce((sum, seg) => sum + seg.text.length, 0) + 1, // +1 for newline
              ...(tabId && { tabId })
            },
            paragraphStyle: {
              namedStyleType: namedStyleType
            },
            fields: 'namedStyleType'
          }
        });
      }
      
      // Apply text styling if the segment has formatting
      if (segment.styles && Object.keys(segment.styles).length > 0) {
        const apiTextStyle: any = {};
        
        if (segment.styles.bold !== undefined) apiTextStyle.bold = segment.styles.bold;
        if (segment.styles.italic !== undefined) apiTextStyle.italic = segment.styles.italic;
        if (segment.styles.strikethrough !== undefined) apiTextStyle.strikethrough = segment.styles.strikethrough;
        if (segment.styles.code) {
          // Code formatting: monospace font
          apiTextStyle.weightedFontFamily = {
            fontFamily: 'Consolas',
            weight: 400
          };
          apiTextStyle.backgroundColor = {
            color: {
              rgbColor: { red: 0.96, green: 0.96, blue: 0.96 }
            }
          };
        }
        
        requests.push({
          updateTextStyle: {
            range: {
              startIndex: currentIndex,
              endIndex: currentIndex + segment.text.length, // Don't include newline in text style
              ...(tabId && { tabId })
            },
            textStyle: apiTextStyle,
            fields: Object.keys(apiTextStyle).join(',')
          }
        });
      }
      
      currentIndex = segmentEndIndex;
    }
  }
  
  return requests;
}

/**
 * Parse inline markdown formatting within a text string
 * Returns an array of text segments with their associated styles
 */
function parseInlineFormatting(text: string): Array<{text: string, styles?: any}> {
  const segments: Array<{text: string, styles?: any}> = [];
  let currentPos = 0;
  
  // Define regex patterns for inline formatting
  const patterns = [
    { regex: /\*\*([^*]+?)\*\*/g, style: { bold: true } },           // **bold**
    { regex: /(?<!\*)\*([^*]+?)\*(?!\*)/g, style: { italic: true } }, // *italic* (not part of **)
    { regex: /~~([^~]+?)~~/g, style: { strikethrough: true } },      // ~~strikethrough~~
    { regex: /`([^`]+?)`/g, style: { code: true } }                  // `code`
  ];
  
  // Find all matches and their positions
  const matches: Array<{start: number, end: number, text: string, style: any}> = [];
  
  for (const pattern of patterns) {
    let match;
    const regex = new RegExp(pattern.regex.source, pattern.regex.flags);
    
    while ((match = regex.exec(text)) !== null) {
      matches.push({
        start: match.index,
        end: match.index + match[0].length,
        text: match[1], // The captured group (text without markup)
        style: pattern.style
      });
    }
  }
  
  // Sort matches by start position
  matches.sort((a, b) => a.start - b.start);
  
  // Handle overlapping matches (give priority to the first one found)
  const validMatches = [];
  let lastEnd = 0;
  
  for (const match of matches) {
    if (match.start >= lastEnd) {
      validMatches.push(match);
      lastEnd = match.end;
    }
  }
  
  // Build segments
  let pos = 0;
  
  for (const match of validMatches) {
    // Add plain text before the match
    if (match.start > pos) {
      const plainText = text.substring(pos, match.start);
      if (plainText) {
        segments.push({ text: plainText });
      }
    }
    
    // Add the styled text
    segments.push({ 
      text: match.text, 
      styles: match.style 
    });
    
    pos = match.end;
  }
  
  // Add remaining plain text
  if (pos < text.length) {
    const plainText = text.substring(pos);
    if (plainText) {
      segments.push({ text: plainText });
    }
  }
  
  // If no matches found, return the original text as a single segment
  if (segments.length === 0) {
    segments.push({ text: text });
  }
  
  return segments;
}

/**
 * Google Docs Writer Component
 * Clears existing content and writes the new formatted content
 */
async function writeMarkdownToGoogleDoc(docId: string, markdownString: string, tabId?: string): Promise<string> {
  try {
    const documentId = docId.toString();
    
    // Get the document to understand its structure
    const doc = await (docsClient.documents.get as any)({
      documentId,
      includeTabsContent: true,
    });
    
    let targetBody = doc.data.body;
    let baseIndex = 1;
    
    // If tabId is specified, find the target tab
    if (tabId && doc.data.tabs) {
      const allTabs = collectAllTabs(doc.data.tabs);
      const targetTab = allTabs.find(tab => 
        tab.tabProperties?.tabId === tabId || 
        (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
      );
      
      if (targetTab && targetTab.documentTab) {
        targetBody = targetTab.documentTab.body;
        baseIndex = 1;
      } else {
        throw new Error(`Tab "${tabId}" not found`);
      }
    }
    
    // Calculate document length
    let documentLength = baseIndex;
    if (targetBody && targetBody.content && targetBody.content.length > 0) {
      const textContent = extractTextFromContent(targetBody.content);
      documentLength = baseIndex + Math.max(0, textContent.length - 1);
    }
    
    // Step 1: Clear existing content (except title) in a separate operation
    if (documentLength > baseIndex) {
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: [{
            deleteContentRange: {
              range: {
                startIndex: baseIndex,
                endIndex: documentLength,
                ...(tabId && { tabId }),
              },
            },
          }],
        },
      });
    }
    
    // Step 2: Insert new content with formatting in a separate operation
    // Parse markdown and create requests
    const markdownRequests = parseMarkdownToDocRequests(markdownString, baseIndex, tabId);
    
    // Add tab context to all requests if needed
    if (tabId) {
      markdownRequests.forEach(request => {
        if (request.insertText && request.insertText.location) {
          request.insertText.location.tabId = tabId;
        }
        if (request.updateParagraphStyle && request.updateParagraphStyle.range) {
          request.updateParagraphStyle.range.tabId = tabId;
        }
        if (request.updateTextStyle && request.updateTextStyle.range) {
          request.updateTextStyle.range.tabId = tabId;
        }
      });
    }
    
    // Execute markdown insertion and formatting requests
    if (markdownRequests.length > 0) {
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: markdownRequests,
        },
      });
    }
    
    return `Successfully updated document with ${Math.floor(markdownRequests.length / 2)} elements`;
  } catch (error) {
    if (error instanceof Error) {
      console.error("Full error:", error.stack || error.message);
    } else {
      console.error("Non-standard error:", error);
    }
    throw new Error(`Failed to write markdown to document: ${error instanceof Error ? error.message : String(error)}`);
  }
}

// =============================================
// NEW MCP TOOLS BASED ON REQUIREMENTS
// =============================================

// Tool 1: Read document as markdown
server.tool(
  "read-doc-as-markdown",
  {
    docId: z.string().describe("The ID of the document to read as markdown"),
    tabId: z.string().optional().describe("Tab ID to read from (for tabbed documents)"),
  },
  async ({ docId, tabId }) => {
    try {
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
      const documentId = docId.toString();
      
      // Get the document content
      const doc = await (docsClient.documents.get as any)({
        documentId,
        includeTabsContent: true,
      });
      
      let targetContent = doc.data;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetContent = { body: targetTab.documentTab.body };
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      // Convert document to markdown
      const markdownContent = convertDocToMarkdown(targetContent);
      
      const tabInfo = tabId ? ` from tab "${tabId}"` : '';
      
      return {
        content: [
          {
            type: "text",
            text: `Document converted to markdown successfully${tabInfo}!\n\n--- MARKDOWN CONTENT ---\n${markdownContent}\n--- END MARKDOWN ---`,
          },
        ],
      };
    } catch (error) {
      console.error("Error reading document as markdown:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error reading document as markdown: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);



// Tool 2: Write markdown to document
server.tool(
  "write-markdown-to-doc",
  {
    docId: z.string().describe("The ID of the document to update"),
    markdownString: z.string().describe("The markdown content to write to the document"),
    tabId: z.string().optional().describe("Tab ID to write to (for tabbed documents)"),
  },
  async ({ docId, markdownString, tabId }) => {
    try {
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
      // Write the markdown content to the document
      const result = await writeMarkdownToGoogleDoc(docId, markdownString, tabId);
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      
      return {
        content: [
          {
            type: "text",
            text: `${result}${tabInfo}!\nDocument ID: ${docId}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error writing markdown to document:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error writing markdown to document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);


// Connect to the transport and start the server
async function main() {
  // Create a transport for communicating over stdin/stdout
  const transport = new StdioServerTransport();

  // Connect the server to the transport
  await server.connect(transport);
  
  console.error("Google Docs MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});