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
          tab.tabId === tabId || 
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
      let documentLength = baseIndex;
      const traverseContent = (content: any[], currentIndex: number = baseIndex): number => {
        content.forEach((element: any) => {
          if (element.paragraph) {
            element.paragraph.elements.forEach((paragraphElement: any) => {
              if (paragraphElement.textRun && paragraphElement.textRun.content) {
                currentIndex += paragraphElement.textRun.content.length;
              }
            });
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
        documentLength = traverseContent(targetBody.content, baseIndex);
      }
      
      const requests: any[] = [];
      let insertIndex: number;
      
      if (replaceAll) {
        // Delete all content first
        requests.push({
          deleteContentRange: {
            range: {
              startIndex: baseIndex,
              endIndex: documentLength,
            },
          },
        });
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
        
        requests.push({
          updateTextStyle: {
            range: {
              startIndex: insertIndex,
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            textStyle: textStyle,
            fields: Object.keys(textStyle).join(','),
          },
        });
      }
      
      // Apply paragraph styling if specified
      if (paragraphStyle && Object.keys(paragraphStyle).length > 0) {
        const endIndex = insertIndex + content.length;
        
        requests.push({
          updateParagraphStyle: {
            range: {
              startIndex: insertIndex,
              endIndex: endIndex,
              ...(tabId && { tabId }),
            },
            paragraphStyle: paragraphStyle,
            fields: Object.keys(paragraphStyle).join(','),
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
        
        // Calculate the document length
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          doc.data.body.content.forEach((element: any) => {
            if (element.paragraph) {
              element.paragraph.elements.forEach((paragraphElement: any) => {
                if (paragraphElement.textRun && paragraphElement.textRun.content) {
                  documentLength += paragraphElement.textRun.content.length;
                }
              });
            }
          });
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
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          doc.data.body.content.forEach((element: any) => {
            if (element.paragraph) {
              element.paragraph.elements.forEach((paragraphElement: any) => {
                if (paragraphElement.textRun && paragraphElement.textRun.content) {
                  documentLength += paragraphElement.textRun.content.length;
                }
              });
            }
          });
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
          tab.tabId === tabId || 
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

// Tool to update document with formatting preserved/specified
server.tool(
  "update-doc-with-style",
  {
    docId: z.string().describe("The ID of the document to update"),
    content: z.string().describe("The new content to add"),
    replaceAll: z.boolean().optional().describe("Whether to replace all content (true) or append (false). Default: false"),
    insertionPoint: z.number().optional().describe("Specific index to insert at (1-based). If not provided, appends to end"),
    tabId: z.string().optional().describe("Tab ID to insert into (for tabbed documents)"),
    preserveFormatting: z.boolean().optional().describe("Whether to preserve formatting at insertion point. Default: true"),
    textStyle: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name (e.g., 'Arial', 'Times New Roman')"),
      foregroundColor: z.object({
        red: z.number().min(0).max(1).optional(),
        green: z.number().min(0).max(1).optional(),
        blue: z.number().min(0).max(1).optional(),
      }).optional().describe("Text color as RGB values (0-1)"),
      backgroundColor: z.object({
        red: z.number().min(0).max(1).optional(),
        green: z.number().min(0).max(1).optional(),
        blue: z.number().min(0).max(1).optional(),
      }).optional().describe("Background color as RGB values (0-1)"),
    }).optional().describe("Text style to apply. Only specify properties you want to change"),
    paragraphStyle: z.object({
      alignment: z.enum(['ALIGNMENT_UNSPECIFIED', 'START', 'CENTER', 'END', 'JUSTIFIED']).optional(),
      lineSpacing: z.number().optional().describe("Line spacing (e.g., 1.0 = single, 1.5 = 1.5x, 2.0 = double)"),
      spaceAbove: z.number().optional().describe("Space above paragraph in points"),
      spaceBelow: z.number().optional().describe("Space below paragraph in points"),
    }).optional().describe("Paragraph style to apply"),
  },
  async ({ docId, content, replaceAll = false, insertionPoint, tabId, preserveFormatting = true, textStyle, paragraphStyle }) => {
    try {
      // First get the document to understand its structure and current formatting
      const docParams: any = {
        documentId: docId,
      };
      
      // Add tabs content parameter with type assertion
      if (tabId) {
        (docParams as any).includeTabsContent = true;
      }

      const doc: any = await docsClient.documents.get(docParams);
      
      let requests: any[] = [];
      let targetIndex = 1; // Default to beginning
      let targetTabId = tabId;

      // Handle tabbed documents
      if (doc.data.tabs && doc.data.tabs.length > 0) {
        if (!targetTabId) {
          targetTabId = doc.data.tabs[0].tabProperties?.tabId;
        }
      }

      if (replaceAll) {
        // Replace all content
        const range: any = {
          startIndex: 1,
          endIndex: -1, // Will be set based on content length
        };
        
        if (targetTabId) {
          range.tabId = targetTabId;
        }

        // Get the length of content to replace
        let bodyContent;
        if (doc.data.tabs && doc.data.tabs.length > 0) {
          const targetTab = doc.data.tabs.find((tab: any) => 
            tab.tabProperties?.tabId === targetTabId
          );
          bodyContent = targetTab?.documentTab?.body?.content;
        } else {
          bodyContent = doc.data.body?.content;
        }

        if (bodyContent && bodyContent.length > 0) {
          // Find the last element to get the end index
          const lastElement = bodyContent[bodyContent.length - 1];
          if (lastElement.endIndex) {
            range.endIndex = lastElement.endIndex - 1; // Leave the final newline
          }
        }

        // Delete existing content
        requests.push({
          deleteContentRange: {
            range: range
          }
        });

        targetIndex = 1;
      } else if (insertionPoint) {
        targetIndex = insertionPoint;
      } else {
        // Append to end - find the end of the document
        let bodyContent;
        if (doc.data.tabs && doc.data.tabs.length > 0) {
          const targetTab = doc.data.tabs.find((tab: any) => 
            tab.tabProperties?.tabId === targetTabId
          );
          bodyContent = targetTab?.documentTab?.body?.content;
        } else {
          bodyContent = doc.data.body?.content;
        }

        if (bodyContent && bodyContent.length > 0) {
          const lastElement = bodyContent[bodyContent.length - 1];
          if (lastElement.endIndex) {
            targetIndex = lastElement.endIndex - 1; // Insert before final newline
          }
        }
      }

      // Insert the text
      const insertRequest: any = {
        insertText: {
          text: content,
          location: {
            index: targetIndex,
          }
        }
      };

      if (targetTabId) {
        insertRequest.insertText.location.tabId = targetTabId;
      }

      requests.push(insertRequest);

      // Apply text styling if specified
      if (textStyle || preserveFormatting) {
        const styleRange: any = {
          startIndex: targetIndex,
          endIndex: targetIndex + content.length,
        };

        if (targetTabId) {
          styleRange.tabId = targetTabId;
        }

        let appliedTextStyle: any = {};
        let fields: string[] = [];

        // If preserveFormatting is true and no specific style provided, try to read current style
        if (preserveFormatting && !textStyle && !replaceAll) {
          // Try to get formatting from the character before insertion point
          const prevCharIndex = Math.max(1, targetIndex - 1);
          
          // This is a simplified approach - in a full implementation you'd read the document structure
          // and extract the text style at the previous character position
        }

        // Apply specified text style
        if (textStyle) {
          if (textStyle.bold !== undefined) {
            appliedTextStyle.bold = textStyle.bold;
            fields.push('bold');
          }
          if (textStyle.italic !== undefined) {
            appliedTextStyle.italic = textStyle.italic;
            fields.push('italic');
          }
          if (textStyle.underline !== undefined) {
            appliedTextStyle.underline = textStyle.underline;
            fields.push('underline');
          }
          if (textStyle.strikethrough !== undefined) {
            appliedTextStyle.strikethrough = textStyle.strikethrough;
            fields.push('strikethrough');
          }
          if (textStyle.fontSize) {
            appliedTextStyle.fontSize = {
              magnitude: textStyle.fontSize,
              unit: 'PT'
            };
            fields.push('fontSize');
          }
          if (textStyle.fontFamily) {
            appliedTextStyle.weightedFontFamily = {
              fontFamily: textStyle.fontFamily
            };
            fields.push('weightedFontFamily');
          }
          if (textStyle.foregroundColor) {
            appliedTextStyle.foregroundColor = {
              color: {
                rgbColor: textStyle.foregroundColor
              }
            };
            fields.push('foregroundColor');
          }
          if (textStyle.backgroundColor) {
            appliedTextStyle.backgroundColor = {
              color: {
                rgbColor: textStyle.backgroundColor
              }
            };
            fields.push('backgroundColor');
          }
        }

        if (fields.length > 0) {
          requests.push({
            updateTextStyle: {
              range: styleRange,
              textStyle: appliedTextStyle,
              fields: fields.join(',')
            }
          });
        }
      }

      // Apply paragraph styling if specified
      if (paragraphStyle) {
        const paragraphRange: any = {
          startIndex: targetIndex,
          endIndex: targetIndex + content.length,
        };

        if (targetTabId) {
          paragraphRange.tabId = targetTabId;
        }

        let appliedParagraphStyle: any = {};
        let paragraphFields: string[] = [];

        if (paragraphStyle.alignment) {
          appliedParagraphStyle.alignment = paragraphStyle.alignment;
          paragraphFields.push('alignment');
        }
        if (paragraphStyle.lineSpacing) {
          appliedParagraphStyle.lineSpacing = paragraphStyle.lineSpacing;
          paragraphFields.push('lineSpacing');
        }
        if (paragraphStyle.spaceAbove) {
          appliedParagraphStyle.spaceAbove = {
            magnitude: paragraphStyle.spaceAbove,
            unit: 'PT'
          };
          paragraphFields.push('spaceAbove');
        }
        if (paragraphStyle.spaceBelow) {
          appliedParagraphStyle.spaceBelow = {
            magnitude: paragraphStyle.spaceBelow,
            unit: 'PT'
          };
          paragraphFields.push('spaceBelow');
        }

        if (paragraphFields.length > 0) {
          requests.push({
            updateParagraphStyle: {
              range: paragraphRange,
              paragraphStyle: appliedParagraphStyle,
              fields: paragraphFields.join(',')
            }
          });
        }
      }

      // Execute all requests
      await docsClient.documents.batchUpdate({
        documentId: docId,
        requestBody: {
          requests: requests,
        },
      });

      const actionText = replaceAll ? 'replaced' : 'updated';
      const tabText = targetTabId ? ` in tab ${targetTabId}` : '';
      let styleInfo = '';
      
      if (textStyle || paragraphStyle) {
        const styleDetails = [];
        if (textStyle) styleDetails.push('text formatting');
        if (paragraphStyle) styleDetails.push('paragraph formatting');
        styleInfo = ` with ${styleDetails.join(' and ')}`;
      } else if (preserveFormatting) {
        styleInfo = ' with formatting preserved';
      }

      return {
        content: [
          {
            type: "text",
            text: `Successfully ${actionText} document content${tabText}${styleInfo}. Added ${content.length} characters.`,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating document: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to get text style at a specific location
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
      const docParams: any = {
        documentId: docId,
      };
      
      // Add tabs content parameter if we're working with tabs
      if (tabId) {
        (docParams as any).includeTabsContent = true;
      }

      const doc: any = await docsClient.documents.get(docParams);
      
      let targetTabId = tabId;
      let documentContent;

      // Handle tabbed documents
      if (doc.data.tabs && doc.data.tabs.length > 0) {
        if (!targetTabId) {
          targetTabId = doc.data.tabs[0].tabProperties?.tabId;
        }
        
        const targetTab = doc.data.tabs.find((tab: any) => 
          tab.tabProperties?.tabId === targetTabId
        );
        documentContent = targetTab?.documentTab;
      } else {
        documentContent = doc.data;
      }

      if (!documentContent) {
        return {
          content: [
            {
              type: "text",
              text: `Could not find content${tabId ? ` for tab ${tabId}` : ''}.`,
            },
          ],
          isError: true,
        };
      }

      const actualEndIndex = endIndex || startIndex + 1;
      
      // Find the text style by traversing the document structure
      let foundStyles: any[] = [];
      
      const traverseContent = (content: any[], currentIndex: number = 1): number => {
        for (const element of content) {
          if (element.paragraph) {
            const paragraph = element.paragraph;
            
            if (paragraph.elements) {
              for (const elem of paragraph.elements) {
                if (elem.textRun) {
                  const textRun = elem.textRun;
                  const textLength = textRun.content?.length || 0;
                  const elementEndIndex = currentIndex + textLength;
                  
                  // Check if our target range overlaps with this text run
                  if (currentIndex < actualEndIndex && elementEndIndex > startIndex) {
                    foundStyles.push({
                      range: { startIndex: currentIndex, endIndex: elementEndIndex },
                      textStyle: textRun.textStyle || {},
                      content: textRun.content,
                    });
                  }
                  
                  currentIndex = elementEndIndex;
                } else {
                  // Handle other element types (like page breaks, inline objects)
                  currentIndex += 1;
                }
              }
            }
          } else if (element.table) {
            // Handle tables - traverse table cells
            const table = element.table;
            if (table.tableRows) {
              for (const row of table.tableRows) {
                if (row.tableCells) {
                  for (const cell of row.tableCells) {
                    if (cell.content) {
                      currentIndex = traverseContent(cell.content, currentIndex);
                    }
                  }
                }
              }
            }
          } else if (element.tableOfContents) {
            // Handle table of contents
            currentIndex += 1;
          } else {
            // Handle other structural elements
            currentIndex += 1;
          }
        }
        return currentIndex;
      };

      if (documentContent.body?.content) {
        traverseContent(documentContent.body.content);
      }

      // Format the response
      let responseText = `Text style information for range ${startIndex}`;
      if (endIndex) {
        responseText += `-${endIndex}`;
      }
      responseText += `${tabId ? ` in tab ${tabId}` : ''}:\n\n`;

      if (foundStyles.length === 0) {
        responseText += "No text found in the specified range.";
      } else {
        foundStyles.forEach((style, index) => {
          responseText += `Text segment ${index + 1} (${style.range.startIndex}-${style.range.endIndex}):\n`;
          responseText += `Content: "${style.content?.replace(/\n/g, '\\n')}"\n`;
          
          const textStyle = style.textStyle;
          if (Object.keys(textStyle).length === 0) {
            responseText += "Style: Default (no explicit formatting)\n";
          } else {
            responseText += "Style:\n";
            
            if (textStyle.bold) responseText += `  - Bold: ${textStyle.bold}\n`;
            if (textStyle.italic) responseText += `  - Italic: ${textStyle.italic}\n`;
            if (textStyle.underline) responseText += `  - Underline: ${textStyle.underline}\n`;
            if (textStyle.strikethrough) responseText += `  - Strikethrough: ${textStyle.strikethrough}\n`;
            
            if (textStyle.fontSize) {
              responseText += `  - Font Size: ${textStyle.fontSize.magnitude} ${textStyle.fontSize.unit}\n`;
            }
            
            if (textStyle.weightedFontFamily) {
              responseText += `  - Font Family: ${textStyle.weightedFontFamily.fontFamily}\n`;
            }
            
            if (textStyle.foregroundColor) {
              const color = textStyle.foregroundColor.color?.rgbColor;
              if (color) {
                responseText += `  - Text Color: RGB(${color.red || 0}, ${color.green || 0}, ${color.blue || 0})\n`;
              }
            }
            
            if (textStyle.backgroundColor) {
              const bgColor = textStyle.backgroundColor.color?.rgbColor;
              if (bgColor) {
                responseText += `  - Background Color: RGB(${bgColor.red || 0}, ${bgColor.green || 0}, ${bgColor.blue || 0})\n`;
              }
            }
            
            if (textStyle.link) {
              responseText += `  - Link: ${textStyle.link.url || 'Yes'}\n`;
            }
          }
          
          responseText += "\n";
        });
      }

      return {
        content: [
          {
            type: "text",
            text: responseText,
          },
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `Error reading text style: ${error.message}`,
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