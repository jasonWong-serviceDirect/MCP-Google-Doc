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

// Tool to update an existing document
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