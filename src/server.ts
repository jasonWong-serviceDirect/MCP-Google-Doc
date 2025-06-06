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

// Tool to update an existing document with formatting preservation
server.tool(
  "update-doc-with-style",
  {
    docId: z.string().describe("The ID of the document to update"),
    content: z.string().describe("The new content to add"),
    replaceAll: z.boolean().optional().describe("Whether to replace all content (true) or append (false). Default: false"),
    insertionPoint: z.number().optional().describe("Specific index to insert at (1-based). If not provided, appends to end"),
    preserveFormatting: z.boolean().optional().describe("Whether to preserve formatting at insertion point. Default: true"),
    tabId: z.string().optional().describe("Tab ID to insert into (for tabbed documents)"),
    textStyle: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name (e.g., 'Arial', 'Times New Roman')"),
      foregroundColor: z.object({
        red: z.number().min(0).max(1),
        green: z.number().min(0).max(1),
        blue: z.number().min(0).max(1),
      }).optional().describe("Text color as RGB values (0-1)"),
      backgroundColor: z.object({
        red: z.number().min(0).max(1),
        green: z.number().min(0).max(1),
        blue: z.number().min(0).max(1),
      }).optional().describe("Background color as RGB values (0-1)"),
    }).optional().describe("Text style to apply. Only specify properties you want to change"),
    paragraphStyle: z.object({
      alignment: z.enum(["ALIGNMENT_UNSPECIFIED", "START", "CENTER", "END", "JUSTIFIED"]).optional(),
      lineSpacing: z.number().optional().describe("Line spacing (e.g., 1.0 = single, 1.5 = 1.5x, 2.0 = double)"),
      spaceAbove: z.number().optional().describe("Space above paragraph in points"),
      spaceBelow: z.number().optional().describe("Space below paragraph in points"),
    }).optional().describe("Paragraph style to apply"),
  },
  async ({ 
    docId, 
    content, 
    replaceAll = false, 
    insertionPoint, 
    preserveFormatting = true, 
    tabId, 
    textStyle, 
    paragraphStyle 
  }) => {
    try {
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
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
          baseIndex = 1; // Tabs start at index 1
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      // Calculate document length and insertion point
      // For tabs, we need to use a different approach since tab indexing can be tricky
      let documentLength = baseIndex;
      if (targetBody && targetBody.content && targetBody.content.length > 0) {
        // Extract all text to count actual content length
        const textContent = extractTextFromContent(targetBody.content);
        // Be conservative - use text length from base index, but don't exceed it
        documentLength = baseIndex + Math.max(0, textContent.length - 1);
      }
      
      const requests: any[] = [];
      let insertIndex: number;
      
      if (replaceAll) {
        // For replaceAll, we need to be more careful about the range
        // Only delete if there's actually content to delete
        if (documentLength > baseIndex) {
          requests.push({
            deleteContentRange: {
              range: {
                startIndex: baseIndex,
                endIndex: documentLength,
                ...(tabId && { tabId }),
              },
            },
          });
        }
        insertIndex = baseIndex;
      } else if (insertionPoint) {
        // Use specified insertion point (convert from 1-based to 0-based if needed)
        insertIndex = Math.max(baseIndex, Math.min(insertionPoint, documentLength));
      } else {
        // Append to end
        insertIndex = documentLength;
      }
      
      // Insert the new content
      requests.push({
        insertText: {
          location: {
            index: insertIndex,
            ...(tabId && { tabId }),
          },
          text: content,
        },
      });
      
      // Apply text styling if specified
      if (textStyle && Object.keys(textStyle).length > 0) {
        const endIndex = insertIndex + content.length;
        
        // Convert textStyle to Google Docs API format (using correct camelCase field names)
        const apiTextStyle: any = {};
        
        // Convert fontSize to Dimension object
        if (textStyle.fontSize) {
          apiTextStyle.fontSize = {
            magnitude: textStyle.fontSize,
            unit: "PT"
          };
        }
        
        // Convert fontFamily to weightedFontFamily structure
        if (textStyle.fontFamily) {
          apiTextStyle.weightedFontFamily = {
            fontFamily: textStyle.fontFamily,
            weight: 400 // Default weight
          };
        }
        
        // Convert color properties to correct nested structure
        if (textStyle.foregroundColor) {
          apiTextStyle.foregroundColor = {
            color: {
              rgbColor: textStyle.foregroundColor
            }
          };
        }
        
        if (textStyle.backgroundColor) {
          apiTextStyle.backgroundColor = {
            color: {
              rgbColor: textStyle.backgroundColor
            }
          };
        }
        
        // These properties remain the same
        if (textStyle.bold !== undefined) apiTextStyle.bold = textStyle.bold;
        if (textStyle.italic !== undefined) apiTextStyle.italic = textStyle.italic;
        if (textStyle.underline !== undefined) apiTextStyle.underline = textStyle.underline;
        if (textStyle.strikethrough !== undefined) apiTextStyle.strikethrough = textStyle.strikethrough;
        
        requests.push({
          updateTextStyle: {
            range: {
              startIndex: insertIndex,
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            textStyle: apiTextStyle,
            fields: Object.keys(apiTextStyle).join(','),
          },
        });
      }
      
      // Apply paragraph styling if specified
      if (paragraphStyle && Object.keys(paragraphStyle).length > 0) {
        const endIndex = insertIndex + content.length;
        
        // Convert paragraphStyle to Google Docs API format (using correct camelCase field names)
        const apiParagraphStyle: any = {};
        
        if (paragraphStyle.alignment) {
          apiParagraphStyle.alignment = paragraphStyle.alignment;
        }
        
        if (paragraphStyle.lineSpacing) {
          apiParagraphStyle.lineSpacing = paragraphStyle.lineSpacing;
        }
        
        if (paragraphStyle.spaceAbove) {
          apiParagraphStyle.spaceAbove = {
            magnitude: paragraphStyle.spaceAbove,
            unit: "PT"
          };
        }
        
        if (paragraphStyle.spaceBelow) {
          apiParagraphStyle.spaceBelow = {
            magnitude: paragraphStyle.spaceBelow,
            unit: "PT"
          };
        }
        
        requests.push({
          updateParagraphStyle: {
            range: {
              startIndex: insertIndex,
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            paragraphStyle: apiParagraphStyle,
            fields: Object.keys(apiParagraphStyle).join(','),
          },
        });
      }
      
      // Execute all requests
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests,
        },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Document updated successfully with formatting preserved!${tabInfo}\nDocument ID: ${docId}\nContent inserted at index: ${insertIndex}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error updating document with style:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Keep the original update-doc for backward compatibility
server.tool(
  "update-doc",
  {
    docId: z.string().describe("The ID of the document to update"),
    content: z.string().describe("The content to add to the document"),
    replaceAll: z.boolean().optional().describe("Whether to replace all content (true) or append (false)"),
  },
  async ({ docId, content, replaceAll = false }) => {
    try {
      // Ensure docId is a string and not null/undefined
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
      const documentId = docId.toString();
      
      if (replaceAll) {
        // First, get the document to find its length
        const doc = await docsClient.documents.get({
          documentId,
        });
        
        // Calculate the document length more accurately
        // Google Docs API provides endIndex for each structural element
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          // Find the last structural element and use its endIndex - 1
          const lastElement = doc.data.body.content[doc.data.body.content.length - 1];
          if (lastElement && lastElement.endIndex) {
            documentLength = lastElement.endIndex - 1; // Delete up to but not including the final newline
          } else {
            // Fallback to text extraction method
            const textContent = extractTextFromContent(doc.data.body.content);
            documentLength = Math.max(1, textContent.length);
          }
        }
        
        // Delete all content and then insert new content
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                deleteContentRange: {
                  range: {
                    startIndex: 1,
                    endIndex: documentLength,
                  },
                },
              },
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
      } else {
        // Append content to the end of the document
        const doc = await docsClient.documents.get({
          documentId,
        });
        
        // Calculate the document length to append at the end
        // Google Docs API provides endIndex for each structural element
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          // Find the last structural element and use its endIndex - 1
          const lastElement = doc.data.body.content[doc.data.body.content.length - 1];
          if (lastElement && lastElement.endIndex) {
            documentLength = lastElement.endIndex - 1; // Insert before the final newline
          } else {
            // Fallback to text extraction method
            const textContent = extractTextFromContent(doc.data.body.content);
            documentLength = Math.max(1, textContent.length);
          }
        }
        
        // Append content at the end
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: {
                    index: documentLength,
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
            text: `Document updated successfully!\nDocument ID: ${docId}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error updating document:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to get text style from a specific location in document
server.tool(
  "get-text-style",
  {
    docId: z.string().describe("The ID of the document to read"),
    startIndex: z.number().describe("Start index to get style from (1-based)"),
    endIndex: z.number().optional().describe("End index (1-based). If not provided, will get style of single character"),
    tabId: z.string().optional().describe("Tab ID to read from (for tabbed documents)"),
  },
  async ({ docId, startIndex, endIndex, tabId }) => {
    try {
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
      const documentId = docId.toString();
      
      // Get the document with full content
      const doc = await (docsClient.documents.get as any)({
        documentId,
        includeTabsContent: true,
      });
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      const effectiveEndIndex = endIndex || startIndex + 1;
      const styleInfo: any = {
        range: {
          startIndex,
          endIndex: effectiveEndIndex,
        },
        textStyles: [],
        paragraphStyles: [],
      };
      
      // Traverse content to find styles at the specified range
      const traverseContent = (content: any[], currentIndex: number = 1): number => {
        content.forEach((element: any) => {
          if (element.paragraph) {
            // Check paragraph style
            if (currentIndex <= effectiveEndIndex && currentIndex + (element.paragraph.elements?.length || 0) >= startIndex) {
              if (element.paragraph.paragraphStyle) {
                styleInfo.paragraphStyles.push({
                  range: { startIndex: currentIndex, endIndex: currentIndex + (element.paragraph.elements?.length || 0) },
                  style: element.paragraph.paragraphStyle,
                });
              }
              
              // Check text styles within paragraph
              element.paragraph.elements.forEach((paragraphElement: any) => {
                if (paragraphElement.textRun) {
                  const textLength = paragraphElement.textRun.content?.length || 0;
                  const textEndIndex = currentIndex + textLength;
                  
                  if (currentIndex <= effectiveEndIndex && textEndIndex >= startIndex) {
                    styleInfo.textStyles.push({
                      range: { startIndex: currentIndex, endIndex: textEndIndex },
                      content: paragraphElement.textRun.content,
                      style: paragraphElement.textRun.textStyle || {},
                    });
                  }
                  
                  currentIndex += textLength;
                }
              });
            }
          } else if (element.table) {
            element.table.tableRows.forEach((row: any) => {
              row.tableCells.forEach((cell: any) => {
                if (cell.content) {
                  currentIndex = traverseContent(cell.content, currentIndex);
                }
              });
            });
          }
        });
        return currentIndex;
      };
      
      if (targetBody && targetBody.content) {
        traverseContent(targetBody.content, 1);
      }
      
      const tabInfo = tabId ? ` from tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Text style information${tabInfo}:\n\n${JSON.stringify(styleInfo, null, 2)}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error getting text style:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error getting text style: ${error}`,
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

// Tool to read a document
server.tool(
  "read-doc",
  {
    docId: z.string().describe("The ID of the document to read"),
  },
  async ({ docId }) => {
    try {
      const doc = await docsClient.documents.get({
        documentId: docId,
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
        content: [
          {
            type: "text",
            text: content,
          },
        ],
      };
    } catch (error) {
      console.error(`Error reading document ${docId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error reading document: ${error}`,
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

// Tool to read from a specific tab by name
server.tool(
  "read-doc-tab",
  {
    docId: z.string().describe("The ID of the document to read"),
    tabName: z.string().describe("The name of the tab to read from"),
  },
  async ({ docId, tabName }) => {
    try {
      // Use the googleapis client but bypass TypeScript checking for the new parameter
      // The includeTabsContent parameter is officially supported by the Google Docs API
      // Reference: https://developers.google.com/workspace/docs/api/how-tos/tabs
      const response = await (docsClient.documents as any).get({
        documentId: docId,
        includeTabsContent: true,
      });
      
      const doc = response.data;
      let content = `Document: ${doc.title}\n`;
      
      // Check if document has tabs
      if (!doc.tabs || doc.tabs.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `Document "${doc.title}" does not have any tabs. Use the regular read-doc tool instead.`,
            },
          ],
          isError: true,
        };
      }
      
      // Collect all tabs (including child tabs)
      const allTabs = collectAllTabs(doc.tabs);
      
      // Find the tab with the specified name (case-insensitive)
      const targetTab = allTabs.find((tab: any) => {
        const tabTitle = tab.tabProperties?.title;
        return tabTitle && tabTitle.toLowerCase().trim() === tabName.toLowerCase().trim();
      });
      
      if (!targetTab) {
        // List available tabs for user reference
        const availableTabs = allTabs.map((tab: any) => 
          tab.tabProperties?.title || "Untitled Tab"
        );
        
        return {
          content: [
            {
              type: "text",
              text: `Tab "${tabName}" not found in document "${doc.title}".\n\nAvailable tabs:\n${availableTabs.map(name => `- ${name}`).join('\n')}`,
            },
          ],
          isError: true,
        };
      }
      
      // Extract content from the found tab
      const tabTitle = targetTab.tabProperties?.title || "Untitled Tab";
      content += `Tab: ${tabTitle}\n\n`;
      
      if (targetTab.documentTab?.body?.content) {
        const textContent = extractTextFromContent(targetTab.documentTab.body.content);
        content += textContent;
      } else {
        content += "No content found in this tab.";
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
      console.error(`Error reading tab "${tabName}" from document ${docId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error reading tab "${tabName}" from document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

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
// HEADING-BASED AND SECTION-AWARE TOOLS
// =============================================

// Tool to find all headings in a document and their positions
server.tool(
  "find-headings",
  {
    docId: z.string().describe("The ID of the document to analyze"),
    tabId: z.string().optional().describe("Tab ID to analyze (for tabbed documents)"),
  },
  async ({ docId, tabId }) => {
    try {
      const doc = await (docsClient.documents.get as any)({
        documentId: docId,
        includeTabsContent: true,
      });
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      const headings = findHeadingsWithPositions(targetBody?.content || []);
      
      let content = `Headings found in document${tabId ? ` (tab: ${tabId})` : ""}:\n\n`;
      
      if (headings.length === 0) {
        content += "No headings found in this document.";
      } else {
        headings.forEach((heading, index) => {
          content += `${index + 1}. ${heading.level}: "${heading.text}" (Start: ${heading.startIndex}, End: ${heading.endIndex})\n`;
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
      console.error("Error finding headings:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error finding headings: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to insert content after a specific heading
server.tool(
  "insert-content-after-heading",
  {
    docId: z.string().describe("The ID of the document"),
    headingText: z.string().describe("The text of the heading to insert after"),
    content: z.string().describe("The content to insert"),
    textStyle: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name"),
      foregroundColor: z.object({
        red: z.number().min(0).max(1),
        green: z.number().min(0).max(1),
        blue: z.number().min(0).max(1),
      }).optional().describe("Text color as RGB values (0-1)"),
    }).optional().describe("Text style to apply"),
    tabId: z.string().optional().describe("Tab ID (for tabbed documents)"),
  },
  async ({ docId, headingText, content, textStyle, tabId }) => {
    try {
      const doc = await (docsClient.documents.get as any)({
        documentId: docId,
        includeTabsContent: true,
      });
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      const headings = findHeadingsWithPositions(targetBody?.content || []);
      const targetHeading = headings.find(h => 
        h.text.toLowerCase().trim() === headingText.toLowerCase().trim()
      );
      
      if (!targetHeading) {
        throw new Error(`Heading "${headingText}" not found`);
      }
      
      const insertIndex = targetHeading.endIndex;
      const requests: any[] = [];
      
      // Insert the content
      requests.push({
        insertText: {
          location: {
            index: insertIndex,
            ...(tabId && { tabId }),
          },
          text: content,
        },
      });
      
      // Apply text styling if specified
      if (textStyle && Object.keys(textStyle).length > 0) {
        const endIndex = insertIndex + content.length;
        const apiTextStyle: any = {};
        
        if (textStyle.fontSize) {
          apiTextStyle.fontSize = { magnitude: textStyle.fontSize, unit: "PT" };
        }
        if (textStyle.fontFamily) {
          apiTextStyle.weightedFontFamily = { fontFamily: textStyle.fontFamily, weight: 400 };
        }
        if (textStyle.foregroundColor) {
          apiTextStyle.foregroundColor = { color: { rgbColor: textStyle.foregroundColor } };
        }
        if (textStyle.bold !== undefined) apiTextStyle.bold = textStyle.bold;
        if (textStyle.italic !== undefined) apiTextStyle.italic = textStyle.italic;
        if (textStyle.underline !== undefined) apiTextStyle.underline = textStyle.underline;
        if (textStyle.strikethrough !== undefined) apiTextStyle.strikethrough = textStyle.strikethrough;
        
        requests.push({
          updateTextStyle: {
            range: { 
              startIndex: insertIndex, 
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            textStyle: apiTextStyle,
            fields: Object.keys(apiTextStyle).join(','),
          },
        });
      }
      
      await docsClient.documents.batchUpdate({
        documentId: docId,
        requestBody: { requests },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Content inserted after heading "${headingText}" successfully${tabInfo} at index ${insertIndex}.`,
          },
        ],
      };
    } catch (error) {
      console.error("Error inserting content after heading:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error inserting content after heading: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to replace content in a section (between two headings)
server.tool(
  "replace-section-content",
  {
    docId: z.string().describe("The ID of the document"),
    startHeading: z.string().describe("The heading that starts the section"),
    endHeading: z.string().optional().describe("The heading that ends the section (if not provided, replaces until next heading or end of document)"),
    newContent: z.string().describe("The new content for the section"),
    preserveHeading: z.boolean().optional().describe("Whether to keep the start heading. Default: true"),
    textStyle: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name"),
      foregroundColor: z.object({
        red: z.number().min(0).max(1),
        green: z.number().min(0).max(1),
        blue: z.number().min(0).max(1),
      }).optional().describe("Text color as RGB values (0-1)"),
    }).optional().describe("Text style to apply"),
    tabId: z.string().optional().describe("Tab ID (for tabbed documents)"),
  },
  async ({ docId, startHeading, endHeading, newContent, preserveHeading = true, textStyle, tabId }) => {
    try {
      const doc = await (docsClient.documents.get as any)({
        documentId: docId,
        includeTabsContent: true,
      });
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      const headings = findHeadingsWithPositions(targetBody?.content || []);
      const startHeadingObj = headings.find(h => 
        h.text.toLowerCase().trim() === startHeading.toLowerCase().trim()
      );
      
      if (!startHeadingObj) {
        throw new Error(`Start heading "${startHeading}" not found`);
      }
      
      let sectionStartIndex = preserveHeading ? startHeadingObj.endIndex : startHeadingObj.startIndex;
      let sectionEndIndex: number;
      
      if (endHeading) {
        const endHeadingObj = headings.find(h => 
          h.text.toLowerCase().trim() === endHeading.toLowerCase().trim()
        );
        if (!endHeadingObj) {
          throw new Error(`End heading "${endHeading}" not found`);
        }
        sectionEndIndex = endHeadingObj.startIndex;
      } else {
        // Find next heading or end of document
        const nextHeading = headings.find(h => h.startIndex > startHeadingObj.endIndex);
        if (nextHeading) {
          sectionEndIndex = nextHeading.startIndex;
        } else {
          // Get document length
          const textContent = extractTextFromContent(targetBody?.content || []);
          sectionEndIndex = Math.max(1, textContent.length);
        }
      }
      
      const requests: any[] = [];
      
      // Delete the section content
      if (sectionEndIndex > sectionStartIndex) {
        requests.push({
          deleteContentRange: {
            range: {
              startIndex: sectionStartIndex,
              endIndex: sectionEndIndex,
              ...(tabId && { tabId }),
            },
          },
        });
      }
      
      // Insert new content
      requests.push({
        insertText: {
          location: {
            index: sectionStartIndex,
            ...(tabId && { tabId }),
          },
          text: newContent,
        },
      });
      
      // Apply text styling if specified
      if (textStyle && Object.keys(textStyle).length > 0) {
        const endIndex = sectionStartIndex + newContent.length;
        const apiTextStyle: any = {};
        
        if (textStyle.fontSize) {
          apiTextStyle.fontSize = { magnitude: textStyle.fontSize, unit: "PT" };
        }
        if (textStyle.fontFamily) {
          apiTextStyle.weightedFontFamily = { fontFamily: textStyle.fontFamily, weight: 400 };
        }
        if (textStyle.foregroundColor) {
          apiTextStyle.foregroundColor = { color: { rgbColor: textStyle.foregroundColor } };
        }
        if (textStyle.bold !== undefined) apiTextStyle.bold = textStyle.bold;
        if (textStyle.italic !== undefined) apiTextStyle.italic = textStyle.italic;
        if (textStyle.underline !== undefined) apiTextStyle.underline = textStyle.underline;
        if (textStyle.strikethrough !== undefined) apiTextStyle.strikethrough = textStyle.strikethrough;
        
        requests.push({
          updateTextStyle: {
            range: { 
              startIndex: sectionStartIndex, 
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            textStyle: apiTextStyle,
            fields: Object.keys(apiTextStyle).join(','),
          },
        });
      }
      
      await docsClient.documents.batchUpdate({
        documentId: docId,
        requestBody: { requests },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Section content replaced successfully${tabInfo}. Section "${startHeading}" updated from index ${sectionStartIndex} to ${sectionStartIndex + newContent.length}.`,
          },
        ],
      };
    } catch (error) {
      console.error("Error replacing section content:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error replacing section content: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to append content to a specific section
server.tool(
  "append-to-section",
  {
    docId: z.string().describe("The ID of the document"),
    sectionHeading: z.string().describe("The heading of the section to append to"),
    content: z.string().describe("The content to append"),
    textStyle: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name"),
      foregroundColor: z.object({
        red: z.number().min(0).max(1),
        green: z.number().min(0).max(1),
        blue: z.number().min(0).max(1),
      }).optional().describe("Text color as RGB values (0-1)"),
    }).optional().describe("Text style to apply"),
    tabId: z.string().optional().describe("Tab ID (for tabbed documents)"),
  },
  async ({ docId, sectionHeading, content, textStyle, tabId }) => {
    try {
      const doc = await (docsClient.documents.get as any)({
        documentId: docId,
        includeTabsContent: true,
      });
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      const headings = findHeadingsWithPositions(targetBody?.content || []);
      const targetHeading = headings.find(h => 
        h.text.toLowerCase().trim() === sectionHeading.toLowerCase().trim()
      );
      
      if (!targetHeading) {
        throw new Error(`Heading "${sectionHeading}" not found`);
      }
      
      // Find the end of this section (before next heading or end of document)
      const nextHeading = headings.find(h => h.startIndex > targetHeading.endIndex);
      let insertIndex: number;
      
      if (nextHeading) {
        insertIndex = nextHeading.startIndex;
      } else {
        // End of document
        const textContent = extractTextFromContent(targetBody?.content || []);
        insertIndex = Math.max(1, textContent.length);
      }
      
      const requests: any[] = [];
      
      // Insert the content
      requests.push({
        insertText: {
          location: {
            index: insertIndex,
            ...(tabId && { tabId }),
          },
          text: content,
        },
      });
      
      // Apply text styling if specified
      if (textStyle && Object.keys(textStyle).length > 0) {
        const endIndex = insertIndex + content.length;
        const apiTextStyle: any = {};
        
        if (textStyle.fontSize) {
          apiTextStyle.fontSize = { magnitude: textStyle.fontSize, unit: "PT" };
        }
        if (textStyle.fontFamily) {
          apiTextStyle.weightedFontFamily = { fontFamily: textStyle.fontFamily, weight: 400 };
        }
        if (textStyle.foregroundColor) {
          apiTextStyle.foregroundColor = { color: { rgbColor: textStyle.foregroundColor } };
        }
        if (textStyle.bold !== undefined) apiTextStyle.bold = textStyle.bold;
        if (textStyle.italic !== undefined) apiTextStyle.italic = textStyle.italic;
        if (textStyle.underline !== undefined) apiTextStyle.underline = textStyle.underline;
        if (textStyle.strikethrough !== undefined) apiTextStyle.strikethrough = textStyle.strikethrough;
        
        requests.push({
          updateTextStyle: {
            range: { 
              startIndex: insertIndex, 
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            textStyle: apiTextStyle,
            fields: Object.keys(apiTextStyle).join(','),
          },
        });
      }
      
      await docsClient.documents.batchUpdate({
        documentId: docId,
        requestBody: { requests },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Content appended to section "${sectionHeading}" successfully${tabInfo} at index ${insertIndex}.`,
          },
        ],
      };
    } catch (error) {
      console.error("Error appending to section:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error appending to section: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// =============================================
// HELPER FUNCTIONS FOR HEADING ANALYSIS
// =============================================

// Helper function to find headings with their exact positions
function findHeadingsWithPositions(content: any[]): any[] {
  const headings = [];
  let currentIndex = 1;
  
  content.forEach((element: any) => {
    if (element.paragraph) {
      const paragraphStyle = element.paragraph.paragraphStyle;
      
      if (paragraphStyle && paragraphStyle.headingId) {
        let headingText = "";
        let paragraphLength = 0;
        
        element.paragraph.elements.forEach((paragraphElement: any) => {
          if (paragraphElement.textRun && paragraphElement.textRun.content) {
            headingText += paragraphElement.textRun.content;
            paragraphLength += paragraphElement.textRun.content.length;
          }
        });
        
        headings.push({
          level: paragraphStyle.headingId,
          text: headingText.trim(),
          startIndex: currentIndex,
          endIndex: currentIndex + paragraphLength,
        });
        
        currentIndex += paragraphLength;
      } else {
        // Regular paragraph, count its length
        element.paragraph.elements.forEach((paragraphElement: any) => {
          if (paragraphElement.textRun && paragraphElement.textRun.content) {
            currentIndex += paragraphElement.textRun.content.length;
          }
        });
      }
    } else if (element.table) {
      // Count table content length
      element.table.tableRows.forEach((row: any) => {
        row.tableCells.forEach((cell: any) => {
          if (cell.content) {
            const cellText = extractTextFromContent(cell.content);
            currentIndex += cellText.length;
          }
        });
      });
    }
  });
  
  return headings;
}

// =============================================
// MARKDOWN PARSING AND CONVERSION TOOLS
// =============================================

// Helper function to parse markdown content and identify elements
function parseMarkdownContent(markdownText: string): any[] {
  const lines = markdownText.split('\n');
  const elements = [];
  
  for (const line of lines) {
    // Check for headings (# ## ### #### ##### ######)
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const text = headingMatch[2].trim();
      elements.push({
        type: 'heading',
        level: level,
        text: text + '\n',
        namedStyleType: `HEADING_${level}`
      });
    } else {
      // Regular paragraph
      if (line.trim() || elements.length === 0) {
        elements.push({
          type: 'paragraph',
          text: line + '\n',
          namedStyleType: 'NORMAL_TEXT'
        });
      }
    }
  }
  
  return elements;
}

// Tool to insert markdown content and convert to proper Google Docs formatting
server.tool(
  "insert-markdown-content",
  {
    docId: z.string().describe("The ID of the document to update"),
    markdownContent: z.string().describe("The markdown content to insert and convert"),
    insertionPoint: z.number().optional().describe("Specific index to insert at (1-based). If not provided, appends to end"),
    replaceAll: z.boolean().optional().describe("Whether to replace all content (true) or append (false). Default: false"),
    tabId: z.string().optional().describe("Tab ID to insert into (for tabbed documents)"),
  },
  async ({ docId, markdownContent, insertionPoint, replaceAll = false, tabId }) => {
    try {
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
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
      
      // Calculate document length and insertion point
      let documentLength = baseIndex;
      if (targetBody && targetBody.content && targetBody.content.length > 0) {
        const textContent = extractTextFromContent(targetBody.content);
        documentLength = baseIndex + Math.max(0, textContent.length - 1);
      }
      
      // Parse the markdown content
      const elements = parseMarkdownContent(markdownContent);
      
      const requests: any[] = [];
      let insertIndex: number;
      
      if (replaceAll) {
        // For replaceAll, delete existing content first
        if (documentLength > baseIndex) {
          requests.push({
            deleteContentRange: {
              range: {
                startIndex: baseIndex,
                endIndex: documentLength,
                ...(tabId && { tabId }),
              },
            },
          });
        }
        insertIndex = baseIndex;
      } else if (insertionPoint) {
        insertIndex = Math.max(baseIndex, Math.min(insertionPoint, documentLength));
      } else {
        insertIndex = documentLength;
      }
      
      let currentIndex = insertIndex;
      
      // Process each markdown element
      for (const element of elements) {
        // Insert the text
        requests.push({
          insertText: {
            location: {
              index: currentIndex,
              ...(tabId && { tabId }),
            },
            text: element.text,
          },
        });
        
        // Apply the appropriate style based on element type
        const endIndex = currentIndex + element.text.length;
        
        requests.push({
          updateParagraphStyle: {
            range: {
              startIndex: currentIndex,
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            paragraphStyle: {
              namedStyleType: element.namedStyleType,
            },
            fields: 'namedStyleType',
          },
        });
        
        currentIndex = endIndex;
      }
      
      // Execute all requests
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests,
        },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      const elementCount = elements.filter(e => e.type === 'heading').length;
      const headingInfo = elementCount > 0 ? ` (${elementCount} headings converted)` : '';
      
      return {
        content: [
          {
            type: "text",
            text: `Markdown content inserted and converted successfully${tabInfo}${headingInfo}!\nDocument ID: ${docId}\nContent inserted at index: ${insertIndex}\nElements processed: ${elements.length}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error inserting markdown content:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error inserting markdown content: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to convert existing plain text to markdown-formatted content
server.tool(
  "convert-text-to-markdown-headings",
  {
    docId: z.string().describe("The ID of the document to update"),
    startIndex: z.number().describe("Start index of the text range to convert (1-based)"),
    endIndex: z.number().describe("End index of the text range to convert (1-based)"),
    tabId: z.string().optional().describe("Tab ID (for tabbed documents)"),
  },
  async ({ docId, startIndex, endIndex, tabId }) => {
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
      
      let targetBody = doc.data.body;
      
      // If tabId is specified, find the target tab
      if (tabId && doc.data.tabs) {
        const allTabs = collectAllTabs(doc.data.tabs);
        const targetTab = allTabs.find(tab => 
          tab.tabProperties?.tabId === tabId || 
          (tab.documentTab?.title && tab.documentTab.title.toLowerCase() === tabId.toLowerCase())
        );
        
        if (targetTab && targetTab.documentTab) {
          targetBody = targetTab.documentTab.body;
        } else {
          throw new Error(`Tab "${tabId}" not found`);
        }
      }
      
      // Extract text from the specified range
      let extractedText = "";
      let currentIndex = 1;
      
      const extractTextFromRange = (content: any[], start: number, end: number): string => {
        let text = "";
        let index = 1;
        
        content.forEach((element: any) => {
          if (element.paragraph) {
            element.paragraph.elements.forEach((paragraphElement: any) => {
              if (paragraphElement.textRun && paragraphElement.textRun.content) {
                const textContent = paragraphElement.textRun.content;
                const textStart = index;
                const textEnd = index + textContent.length;
                
                // Check if this text overlaps with our target range
                if (textStart < end && textEnd > start) {
                  const extractStart = Math.max(0, start - textStart);
                  const extractEnd = Math.min(textContent.length, end - textStart);
                  text += textContent.substring(extractStart, extractEnd);
                }
                
                index += textContent.length;
              }
            });
          }
        });
        
        return text;
      };
      
      if (targetBody && targetBody.content) {
        extractedText = extractTextFromRange(targetBody.content, startIndex, endIndex);
      }
      
      // Parse the extracted text for markdown patterns
      const elements = parseMarkdownContent(extractedText);
      const headingsFound = elements.filter(e => e.type === 'heading');
      
      if (headingsFound.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No markdown heading patterns found in the specified range (${startIndex}-${endIndex}).`,
            },
          ],
        };
      }
      
      const requests: any[] = [];
      
      // Delete the original text range
      requests.push({
        deleteContentRange: {
          range: {
            startIndex: startIndex,
            endIndex: endIndex,
            ...(tabId && { tabId }),
          },
        },
      });
      
      let currentInsertIndex = startIndex;
      
      // Insert the converted content with proper formatting
      for (const element of elements) {
        // Insert the text
        requests.push({
          insertText: {
            location: {
              index: currentInsertIndex,
              ...(tabId && { tabId }),
            },
            text: element.text,
          },
        });
        
        // Apply the appropriate style
        const elementEndIndex = currentInsertIndex + element.text.length;
        
        requests.push({
          updateParagraphStyle: {
            range: {
              startIndex: currentInsertIndex,
              endIndex: elementEndIndex,
              ...(tabId && { tabId }),
            },
            paragraphStyle: {
              namedStyleType: element.namedStyleType,
            },
            fields: 'namedStyleType',
          },
        });
        
        currentInsertIndex = elementEndIndex;
      }
      
      // Execute all requests
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests,
        },
      });
      
      const tabInfo = tabId ? ` in tab "${tabId}"` : '';
      return {
        content: [
          {
            type: "text",
            text: `Text converted to markdown formatting successfully${tabInfo}!\nDocument ID: ${docId}\nRange: ${startIndex}-${endIndex}\nHeadings converted: ${headingsFound.length}\nTotal elements processed: ${elements.length}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error converting text to markdown headings:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error converting text to markdown headings: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

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