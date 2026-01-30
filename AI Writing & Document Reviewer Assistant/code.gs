/**
 * @OnlyCurrentDoc
 */

// Required OAuth Scopes for native comments
// Add these to appsscript.json or they'll be auto-detected
// ============================================
// Configuration - Set these in Script Properties
// ============================================
// GEMINI_API_KEY: Your Google Gemini API key
// PINECONE_API_KEY: Your Pinecone API key
// PINECONE_HOST: Your Pinecone host
// PINECONE_INDEX: Your Pinecone index name (e.g., company-files)
// RULES_FOLDER_ID: Google Drive folder ID containing writing rules

const CHAT_CONFIG = {
  // Existing properties...
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  PINECONE_API_KEY: PropertiesService.getScriptProperties().getProperty('PINECONE_API_KEY'),
  PINECONE_HOST: PropertiesService.getScriptProperties().getProperty('PINECONE_HOST'),
  PINECONE_INDEX: PropertiesService.getScriptProperties().getProperty('PINECONE_INDEX') || 'document-data',
  RULES_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('RULES_FOLDER_ID'),
  
  // Document extraction settings
  SOURCE_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('SOURCE_FOLDER_ID'),
  OUTPUT_SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('OUTPUT_SPREADSHEET_ID'),
  
  // AI Models
  GEMINI_MODEL: 'gemini-2.5-flash',
  GEMINI_EMBEDDING_MODEL: 'text-embedding-004',
  
  // Other settings...
  INCLUDE_SUBFOLDERS: true,
  MAX_DOCS_PER_RUN: 50,
  AUTO_SYNC_TO_PINECONE: true
};

// ============================================
// Web App Entry Points
// ============================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AI Content Assistant')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = processMessage(data.message, data.conversationHistory);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// Main Chat Function (Conversational)
// ============================================

function processMessage(userMessage, conversationHistory = []) {
  try {
    // Detect if this is a document review request
    const isDocReview = isDocumentReviewRequest(userMessage);
    
    if (isDocReview) {
      // Handle document review
      const docId = extractDocumentId(userMessage);
      if (docId) {
        return reviewDocument(docId, conversationHistory);
      }
    }
    
    // Otherwise, handle as conversational query
    return handleConversationalQuery(userMessage, conversationHistory);
    
  } catch (error) {
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}

function isDocumentReviewRequest(message) {
  const reviewKeywords = [
    'review', 'analyze', 'check', 'feedback', 'docs.google.com',
    'document/d/', 'look at this doc', 'can you review'
  ];
  
  const lowerMessage = message.toLowerCase();
  return reviewKeywords.some(keyword => lowerMessage.includes(keyword));
}

// ============================================
// Conversational Query Handler
// ============================================

function handleConversationalQuery(userMessage, conversationHistory) {
  try {
    // Retrieve relevant context from multiple sources
    const context = gatherContext(userMessage);
    
    // Generate conversational response
    const response = generateConversationalResponse(
      userMessage, 
      context, 
      conversationHistory
    );
    
    return {
      success: true,
      message: response,
      context: context
    };
    
  } catch (error) {
    return {
      success: false,
      message: `Error processing your question: ${error.toString()}`
    };
  }
}

function gatherContext(query) {
  const context = {
    pineconeRules: [],
    driveRules: [],
    relevantDocs: []
  };
  
  try {
    // Get relevant rules from Pinecone
    context.pineconeRules = retrieveFromPinecone(query, 5);
  } catch (e) {
    Logger.log(`Pinecone retrieval error: ${e}`);
  }
  
  try {
    // Search writing rules folder in Google Drive
    context.driveRules = searchRulesInDrive(query);
  } catch (e) {
    Logger.log(`Drive search error: ${e}`);
  }
  
  try {
    // Get list of available rule documents
    context.relevantDocs = listRuleDocuments();
  } catch (e) {
    Logger.log(`Document listing error: ${e}`);
  }
  
  return context;
}

// ============================================
// Google Drive Rules Integration
// ============================================

function searchRulesInDrive(query) {
  if (!CHAT_CONFIG.RULES_FOLDER_ID) {
    return [];
  }
  
  try {
    const folder = DriveApp.getFolderById(CHAT_CONFIG.RULES_FOLDER_ID);
    const files = folder.getFiles();
    const results = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      
      // Only process Google Docs
      if (mimeType === MimeType.GOOGLE_DOCS) {
        try {
          const doc = DocumentApp.openById(file.getId());
          const content = doc.getBody().getText();
          
          // Check if query terms appear in document
          if (isRelevantContent(content, query)) {
            results.push({
              title: file.getName(),
              content: content,
              id: file.getId(),
              url: file.getUrl()
            });
          }
        } catch (e) {
          Logger.log(`Error reading doc ${file.getName()}: ${e}`);
        }
      }
    }
    
    return results;
  } catch (error) {
    Logger.log(`Error searching Drive folder: ${error}`);
    return [];
  }
}

function isRelevantContent(content, query) {
  const queryTerms = query.toLowerCase().split(/\s+/);
  const contentLower = content.toLowerCase();
  
  // Check if at least 2 query terms appear in content
  const matches = queryTerms.filter(term => contentLower.includes(term));
  return matches.length >= Math.min(2, queryTerms.length);
}

function listRuleDocuments() {
  if (!CHAT_CONFIG.RULES_FOLDER_ID) {
    return [];
  }
  
  try {
    const folder = DriveApp.getFolderById(CHAT_CONFIG.RULES_FOLDER_ID);
    const files = folder.getFiles();
    const docs = [];
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        docs.push({
          name: file.getName(),
          id: file.getId(),
          url: file.getUrl()
        });
      }
    }
    
    return docs;
  } catch (error) {
    Logger.log(`Error listing documents: ${error}`);
    return [];
  }
}

function readSpecificRuleDoc(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    return {
      title: doc.getName(),
      content: doc.getBody().getText()
    };
  } catch (error) {
    return null;
  }
}

// ============================================
// Enhanced Document Review with Drive Rules
// ============================================

function reviewDocument(docId, conversationHistory = []) {
  try {
    // Read the document to review
    const docContent = readGoogleDoc(docId);
    if (!docContent.success) {
      return {
        success: false,
        message: formatErrorMessage(docId, docContent.error)
      };
    }

    // Gather rules from multiple sources
    const pineconeRules = retrieveFromPinecone(docContent.text, 10);
    const driveRules = searchRulesInDrive(docContent.text);
    
    // Combine all rules
    const allRules = {
      pinecone: pineconeRules,
      drive: driveRules
    };

    // Analyze document with all available rules
    const analysis = analyzeDocumentWithRules(docContent, allRules, conversationHistory);

    return {
      success: true,
      message: analysis,
      documentTitle: docContent.title,
      rulesUsed: {
        pineconeCount: pineconeRules.length,
        driveDocsCount: driveRules.length
      }
    };

  } catch (error) {
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}

// ============================================
// Gemini AI with Enhanced Context
// ============================================

function generateConversationalResponse(userMessage, context, conversationHistory) {
  const systemPrompt = `You are an AI writing assistant with expertise in content guidelines, style rules, and best practices.

You have access to:
1. Writing rules stored in Pinecone vector database
2. Writing rule documents from Google Drive folder
3. Conversation history for context

Your capabilities:
- Answer questions about writing guidelines and style rules
- Provide examples and explanations from the rule documents
- Review and analyze Google Docs when requested
- Have natural conversations about writing best practices
- Reference specific rules and cite sources

When answering:
- Be conversational, helpful, and detailed
- Cite specific rules when relevant
- Provide examples when helpful
- If you don't have information, say so honestly
- Offer to review documents if the user shares a link`;

  // Build context from retrieved information
  let contextText = '\n\n**AVAILABLE CONTEXT:**\n';
  
  if (context.pineconeRules && context.pineconeRules.length > 0) {
    contextText += '\n**From Knowledge Base (Pinecone):**\n';
    context.pineconeRules.forEach((rule, i) => {
      contextText += `${i + 1}. ${rule.content}\n`;
    });
  }
  
  if (context.driveRules && context.driveRules.length > 0) {
    contextText += '\n**From Writing Rules Documents:**\n';
    context.driveRules.forEach((doc, i) => {
      contextText += `\nDocument: "${doc.title}"\n`;
      contextText += `Content excerpt: ${doc.content.substring(0, 500)}...\n`;
    });
  }
  
  if (context.relevantDocs && context.relevantDocs.length > 0) {
    contextText += '\n**Available Rule Documents:**\n';
    context.relevantDocs.forEach(doc => {
      contextText += `- ${doc.name}\n`;
    });
  }

  const userPrompt = userMessage + contextText;

  // Build conversation
  const messages = [
    { role: 'user', parts: [{ text: systemPrompt }] }
  ];

  if (conversationHistory && conversationHistory.length > 0) {
    conversationHistory.slice(-10).forEach(msg => {
      messages.push({
        role: msg.role === 'user' ? 'user' : 'model',
        parts: [{ text: msg.content }]
      });
    });
  }

  messages.push({
    role: 'user',
    parts: [{ text: userPrompt }]
  });

  return callGemini(messages);
}

function listAvailableGeminiModels() {
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${CHAT_CONFIG.GEMINI_API_KEY}`;
  
  const options = {
    method: 'get',
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    Logger.log(`Response Code: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`Error: ${response.getContentText()}`);
      return;
    }
    
    const result = JSON.parse(response.getContentText());
    
    Logger.log('=== AVAILABLE GEMINI MODELS ===\n');
    
    if (result.models) {
      result.models.forEach(model => {
        if (model.supportedGenerationMethods && 
            model.supportedGenerationMethods.includes('generateContent')) {
          Logger.log(`✓ ${model.name}`);
          Logger.log(`  Display Name: ${model.displayName}`);
          Logger.log(`  Methods: ${model.supportedGenerationMethods.join(', ')}`);
          Logger.log('');
        }
      });
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error listing models: ' + error.toString());
    return null;
  }
}
function analyzeDocumentWithRules(docContent, allRules, conversationHistory) {
  const systemPrompt = `You are an expert content reviewer and editor specializing in maintaining consistent writing style and adherence to company writing guidelines.

You have access to writing rules from TWO sources:
1. Vector database (Pinecone) - semantic search results
2. Google Drive folder - actual rule documents

Your task:
1. Analyze the provided document against ALL available rules
2. Cross-reference rules from both sources
3. Identify deviations and provide specific, actionable suggestions
4. Cite the specific rule and source for each issue

Required response format:
**DOCUMENT ANALYZED:** [Document title]

**RULES SOURCES USED:**
- Pinecone Knowledge Base: [number] rules retrieved
- Google Drive Rule Documents: [number] documents checked

**ISSUES FOUND:**

1. **Location:** [Quote the problematic text]
   **Issue:** [Explain what is wrong and which rule it violates]
   **Suggestion:** [Provide the corrected version]
   **Rule Reference:** [Cite the specific guideline and source]

[Continue for each issue]

**STYLE CONSISTENCY NOTES:**
[Observations about consistency]

**SUMMARY:**
[Brief summary with total number of issues found]

If no issues are found, confirm that the content follows all guidelines.`;

  // Format rules context
  let rulesContext = '\n\n**WRITING RULES FROM PINECONE:**\n';
  if (allRules.pinecone && allRules.pinecone.length > 0) {
    allRules.pinecone.forEach((rule, i) => {
      rulesContext += `${i + 1}. ${rule.content}\n   Source: ${rule.source || 'Knowledge Base'}\n\n`;
    });
  } else {
    rulesContext += 'No rules retrieved from Pinecone.\n\n';
  }

  rulesContext += '**WRITING RULES FROM GOOGLE DRIVE:**\n';
  if (allRules.drive && allRules.drive.length > 0) {
    allRules.drive.forEach((doc, i) => {
      rulesContext += `\nDocument: "${doc.title}"\n`;
      rulesContext += `URL: ${doc.url}\n`;
      rulesContext += `Content:\n${doc.content}\n\n`;
    });
  } else {
    rulesContext += 'No rules found in Google Drive folder.\n\n';
  }

  const userPrompt = `Document Title: ${docContent.title}

Document Content:
${docContent.text}
${rulesContext}

Please analyze this document thoroughly using all available rules.`;

  const messages = [
    { role: 'user', parts: [{ text: systemPrompt }] }
  ];

  if (conversationHistory && conversationHistory.length > 0) {
    conversationHistory.slice(-10).forEach(msg => {
      messages.push({
        role: msg.role === 'user' ? 'user' : 'model',
        parts: [{ text: msg.content }]
      });
    });
  }

  messages.push({
    role: 'user',
    parts: [{ text: userPrompt }]
  });

  return callGemini(messages);
}

// ============================================
// Document Reading Functions
// ============================================

function extractDocumentId(url) {
  const patterns = [
    /\/document\/d\/([a-zA-Z0-9-_]+)/,
    /^([a-zA-Z0-9-_]+)$/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }
  return null;
}

function readGoogleDoc(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const title = doc.getName();
    const body = doc.getBody();
    const text = body.getText();
    
    return {
      success: true,
      title: title,
      text: text,
      documentId: docId
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// Pinecone Integration
// ============================================

function getEmbedding(text) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CHAT_CONFIG.GEMINI_EMBEDDING_MODEL}:embedContent?key=${CHAT_CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    model: `models/${CHAT_CONFIG.GEMINI_EMBEDDING_MODEL}`,
    content: {
      parts: [{
        text: text
      }]
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Embedding Response Code: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`Embedding Error Response: ${responseText}`);
      throw new Error(`Embedding API error (${responseCode}): ${responseText}`);
    }
    
    const result = JSON.parse(responseText);
    
    if (result.embedding && result.embedding.values) {
      return result.embedding.values;
    }
    
    Logger.log(`Unexpected embedding response: ${JSON.stringify(result)}`);
    throw new Error('Failed to generate embedding - unexpected response structure');
    
  } catch (error) {
    Logger.log(`Embedding Error: ${error.toString()}`);
    throw new Error(`Failed to generate embedding: ${error.message}`);
  }
}
function retrieveFromPinecone(text, topK = 5) {
  try {
    const query = text.substring(0, 1000);
    const queryVector = getEmbedding(query);
    return queryPinecone(queryVector, topK);
  } catch (error) {
    Logger.log(`Pinecone retrieval error: ${error}`);
    return [];
  }
}

function queryPinecone(queryVector, topK = 5) {
  // Ensure PINECONE_HOST is properly formatted
  let host = CHAT_CONFIG.PINECONE_HOST;
  
  // Remove any protocol prefix if present
  host = host.replace(/^https?:\/\//, '');
  
  // Build the correct URL
  const url = `https://${host}/query`;
  
  const payload = {
    vector: queryVector,
    topK: topK,
    includeMetadata: true
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Api-Key': CHAT_CONFIG.PINECONE_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`Pinecone error (${responseCode}): ${response.getContentText()}`);
      return [];
    }
    
    const result = JSON.parse(response.getContentText());
    const matches = result.matches || [];
    
    return matches.map(match => ({
      content: match.metadata?.text || match.metadata?.content || '',
      score: match.score,
      source: match.metadata?.source || 'Knowledge Base'
    }));
  } catch (error) {
    Logger.log(`Pinecone query error: ${error.toString()}`);
    return [];
  }
}

// ============================================
// Gemini AI Functions
// ============================================

// ============================================
// Gemini AI Functions (Enhanced Error Handling)
// ============================================

function callGemini(messages) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CHAT_CONFIG.GEMINI_MODEL}:generateContent?key=${CHAT_CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    contents: messages,
    generationConfig: {
      temperature: 0.7,
      topK: 40,
      topP: 0.95,
      maxOutputTokens: 8192
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Log the response for debugging
    Logger.log(`Gemini Response Code: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`Gemini Error Response: ${responseText}`);
      throw new Error(`Gemini API error (${responseCode}): ${responseText}`);
    }
    
    const result = JSON.parse(responseText);
    
    // Check for safety blocks or other issues
    if (result.promptFeedback && result.promptFeedback.blockReason) {
      throw new Error(`Content blocked: ${result.promptFeedback.blockReason}`);
    }
    
    // Extract the response text
    if (result.candidates && 
        result.candidates[0] && 
        result.candidates[0].content && 
        result.candidates[0].content.parts && 
        result.candidates[0].content.parts[0] && 
        result.candidates[0].content.parts[0].text) {
      return result.candidates[0].content.parts[0].text;
    }
    
    // If we get here, the response structure is unexpected
    Logger.log(`Unexpected Gemini response structure: ${JSON.stringify(result)}`);
    throw new Error('Unexpected response structure from Gemini');
    
  } catch (error) {
    Logger.log(`Gemini API Error: ${error.toString()}`);
    throw new Error(`Failed to get response from Gemini: ${error.message}`);
  }
}

// ============================================
// Utility Functions
// ============================================

function formatErrorMessage(docId, error) {
  return `DOCUMENT UNAVAILABLE: ${docId}
REASON: ${error}
ACTION REQUIRED: Please grant access to this document or share it publicly. I cannot analyze content I cannot read.`;
}

// ============================================
// Admin Functions - Upload to Pinecone
// ============================================

function syncDriveFolderToPinecone() {
  if (!CHAT_CONFIG.RULES_FOLDER_ID) {
    Logger.log('RULES_FOLDER_ID not configured');
    return;
  }
  
  const folder = DriveApp.getFolderById(CHAT_CONFIG.RULES_FOLDER_ID);
  const files = folder.getFiles();
  let count = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      try {
        uploadDocumentToPinecone(file.getId(), 'writing-rules');
        count++;
        Logger.log(`Uploaded: ${file.getName()}`);
      } catch (e) {
        Logger.log(`Failed to upload ${file.getName()}: ${e}`);
      }
    }
  }
  
  Logger.log(`Synced ${count} documents to Pinecone`);
}

function uploadDocumentToPinecone(docId, namespace = 'writing-rules') {
  const docContent = readGoogleDoc(docId);
  
  if (!docContent.success) {
    throw new Error(`Failed to read document: ${docContent.error}`);
  }
  
  const chunks = chunkText(docContent.text, 500);
  
  chunks.forEach((chunk, index) => {
    const embedding = getEmbedding(chunk);
    const vectorId = `${docId}_chunk_${index}`;
    
    upsertToPinecone(vectorId, embedding, {
      text: chunk,
      source: docContent.title,
      documentId: docId,
      chunkIndex: index
    }, namespace);
  });
  
  Logger.log(`Uploaded ${chunks.length} chunks from ${docContent.title}`);
}

function chunkText(text, maxLength) {
  const chunks = [];
  const paragraphs = text.split('\n\n');
  let currentChunk = '';
  
  paragraphs.forEach(para => {
    if ((currentChunk + para).length > maxLength && currentChunk) {
      chunks.push(currentChunk.trim());
      currentChunk = para;
    } else {
      currentChunk += (currentChunk ? '\n\n' : '') + para;
    }
  });
  
  if (currentChunk) {
    chunks.push(currentChunk.trim());
  }
  
  return chunks;
}

function upsertToPinecone(id, vector, metadata, namespace = '') {
  // Ensure PINECONE_HOST is properly formatted
  let host = CHAT_CONFIG.PINECONE_HOST;
  host = host.replace(/^https?:\/\//, '');
  
  const url = `https://${host}/vectors/upsert`;
  
  const payload = {
    vectors: [{
      id: id,
      values: vector,
      metadata: metadata
    }],
    namespace: namespace
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Api-Key': CHAT_CONFIG.PINECONE_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`Pinecone upsert error (${responseCode}): ${response.getContentText()}`);
      throw new Error(`Pinecone upsert failed with code ${responseCode}`);
    }
    
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Pinecone upsert error: ${error.toString()}`);
    throw error;
  }
}

// ============================================
// Setup Function
// ============================================

function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  // Uncomment and set your actual values
  // props.setProperty('GEMINI_API_KEY', 'your-gemini-api-key');
  // props.setProperty('PINECONE_API_KEY', 'your-pinecone-api-key');
  // props.setProperty('PINECONE_HOST', 'your-index-xxxxx.svc.environment.pinecone.io');
  // props.setProperty('PINECONE_INDEX', 'company-files');
  // props.setProperty('RULES_FOLDER_ID', 'your-google-drive-folder-id');
  Logger.log('Script properties configured');
}

// Quick fix for Pinecone host
function fixPineconeHost() {
  const props = PropertiesService.getScriptProperties();
  
  // ⚠️ REPLACE THIS with your actual Pinecone host from the console
  // Example: 'company-files-abc123.svc.us-east-1-gcp.pinecone.io'
  const correctHost = 'YOUR-INDEX-NAME-xxxxx.svc.us-east-1-gcp.pinecone.io';
  
  props.setProperty('PINECONE_HOST', correctHost);
  
  Logger.log('✅ PINECONE_HOST updated to: ' + correctHost);
  Logger.log('🔄 Run testPineconeConnection() to verify');
}

function checkAllConfig() {
  Logger.log('=== CURRENT CONFIGURATION ===');
  Logger.log('GEMINI_API_KEY: ' + (CHAT_CONFIG.GEMINI_API_KEY ? '✓ Set (' + CHAT_CONFIG.GEMINI_API_KEY.substring(0, 10) + '...)' : '✗ Missing'));
  Logger.log('PINECONE_API_KEY: ' + (CHAT_CONFIG.PINECONE_API_KEY ? '✓ Set' : '✗ Missing'));
  Logger.log('PINECONE_HOST: ' + (CHAT_CONFIG.PINECONE_HOST || '✗ Missing'));
  Logger.log('PINECONE_INDEX: ' + (CHAT_CONFIG.PINECONE_INDEX || '✗ Missing'));
  Logger.log('RULES_FOLDER_ID: ' + (CHAT_CONFIG.RULES_FOLDER_ID || '✗ Missing'));
  
  if (CHAT_CONFIG.PINECONE_HOST && !CHAT_CONFIG.PINECONE_HOST.includes('.pinecone.io')) {
    Logger.log('\n⚠️  WARNING: PINECONE_HOST looks incorrect!');
    Logger.log('It should end with .pinecone.io');
    Logger.log('Example: company-files-abc123.svc.us-east-1-gcp.pinecone.io');
  }
}

// ============================================
// Test Functions
// ============================================

function testBasicQuery() {
  const result = processMessage("What are the writing rules?", []);
  Logger.log(result);
  return result;
}

function testDocumentReview() {
  const testUrl = 'YOUR_TEST_DOCUMENT_URL'; // Replace with actual doc URL
  const result = processMessage(testUrl, []);
  Logger.log(result);
  return result;
}

function testDriveConnection() {
  if (!CHAT_CONFIG.RULES_FOLDER_ID) {
    Logger.log('❌ RULES_FOLDER_ID not set in Script Properties');
    return;
  }
  
  try {
    const folder = DriveApp.getFolderById(CHAT_CONFIG.RULES_FOLDER_ID);
    Logger.log(`✅ Connected to folder: ${folder.getName()}`);
    
    const files = folder.getFiles();
    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        count++;
        Logger.log(`  📄 ${file.getName()}`);
      }
    }
    Logger.log(`✅ Found ${count} Google Docs in folder`);
  } catch (e) {
    Logger.log(`❌ Error: ${e.toString()}`);
  }
}
// ============================================
// Test Gemini Connection
// ============================================

function testGeminiConnection() {
  try {
    Logger.log('=== GEMINI API TEST ===');
    Logger.log(`API Key configured: ${CHAT_CONFIG.GEMINI_API_KEY ? 'Yes' : 'No'}`);
    Logger.log(`Model: ${CHAT_CONFIG.GEMINI_MODEL}`);
    
    if (!CHAT_CONFIG.GEMINI_API_KEY) {
      Logger.log('❌ GEMINI_API_KEY not set!');
      return;
    }
    
    Logger.log('\nTesting simple query...');
    
    const testMessages = [
      {
        role: 'user',
        parts: [{ text: 'Say "Hello, I am working!" and nothing else.' }]
      }
    ];
    
    const response = callGemini(testMessages);
    Logger.log(`✅ Success! Response: ${response}`);
    
    return response;
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.toString()}`);
    Logger.log('\n📋 Troubleshooting:');
    Logger.log('1. Check if your GEMINI_API_KEY is valid');
    Logger.log('2. Verify the API key has access to Gemini API');
    Logger.log('3. Check if there are any usage quotas exceeded');
    Logger.log('4. Ensure the model name is correct: gemini-2.0-flash-exp');
    return null;
  }
}

function testPineconeConnection() {
  try {
    Logger.log('=== PINECONE CONNECTION TEST ===');
    Logger.log(`Current PINECONE_HOST: ${CHAT_CONFIG.PINECONE_HOST}`);
    Logger.log(`Current PINECONE_INDEX: ${CHAT_CONFIG.PINECONE_INDEX}`);
    
    const testText = "writing rules example";
    Logger.log('\n1. Testing Gemini embedding...');
    const embedding = getEmbedding(testText);
    Logger.log(`✅ Embedding created (dimension: ${embedding.length})`);
    
    Logger.log('\n2. Testing Pinecone query...');
    const results = queryPinecone(embedding, 3);
    Logger.log(`✅ Pinecone returned ${results.length} results`);
    
    if (results.length > 0) {
      Logger.log('\n3. Sample result:');
      Logger.log(JSON.stringify(results[0], null, 2));
    } else {
      Logger.log('\n⚠️  No results returned. This could mean:');
      Logger.log('   - Pinecone index is empty (run syncDriveFolderToPinecone)');
      Logger.log('   - PINECONE_HOST is incorrect');
      Logger.log('   - Namespace mismatch');
    }
  } catch (e) {
    Logger.log(`\n❌ Error: ${e.toString()}`);
    Logger.log('\n📋 Troubleshooting steps:');
    Logger.log('1. Check your Pinecone host format. It should look like:');
    Logger.log('   index-name-xxxxx.svc.us-east-1-gcp.pinecone.io');
    Logger.log('2. Get the correct host from Pinecone console:');
    Logger.log('   - Go to https://app.pinecone.io/');
    Logger.log('   - Click on your index');
    Logger.log('   - Look for "Host" in the Connection details');
    Logger.log('3. Update Script Properties with the correct host');
  }
}

function getPineconeIndexInfo() {
  try {
    let host = CHAT_CONFIG.PINECONE_HOST;
    host = host.replace(/^https?:\/\//, '');
    
    const url = `https://${host}/describe_index_stats`;
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Api-Key': CHAT_CONFIG.PINECONE_API_KEY
      },
      payload: JSON.stringify({}),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    Logger.log('=== PINECONE INDEX INFO ===');
    Logger.log(JSON.stringify(result, null, 2));
    
    return result;
  } catch (e) {
    Logger.log(`❌ Error getting index info: ${e.toString()}`);
    return null;
  }
}

// ============================================
// Analyze Document (Enhanced for Complete Scanning)
// ============================================

function analyzeForHighlighting(docContent, allRules) {
  const systemPrompt = `You are a thorough content reviewer. Analyze the ENTIRE document and find ALL issues.

CRITICAL INSTRUCTIONS:
1. Review the complete document from start to finish
2. Find ALL instances of issues, not just the first few
3. "text" must be EXACT verbatim text from document (15-50 characters)
4. Return up to 50 issues if found
5. Be thorough - check every paragraph

Return ONLY valid JSON:
{
  "issues": [
    {
      "text": "exact text from document",
      "issue": "what's wrong",
      "suggestion": "how to fix",
      "rule": "rule name"
    }
  ],
  "summary": "brief summary including total issues found"
}`;

  let rulesContext = '\n\nWRITING RULES TO CHECK:\n';
  if (allRules.pinecone && allRules.pinecone.length > 0) {
    allRules.pinecone.forEach((rule, i) => {
      rulesContext += `${i + 1}. ${rule.content}\n`;
    });
  }
  
  if (allRules.drive && allRules.drive.length > 0) {
    allRules.drive.forEach((doc) => {
      rulesContext += `\n"${doc.title}": ${doc.content.substring(0, 1000)}\n`;
    });
  }

  // Check document length
  const docLength = docContent.text.length;
  Logger.log(`Document length: ${docLength} characters`);

  let docText = docContent.text;
  
  // If document is very long, analyze in chunks and combine results
  if (docLength > 15000) {
    Logger.log('Long document detected - analyzing in chunks...');
    return analyzeInChunks(docContent, allRules, rulesContext);
  }

  const userPrompt = `DOCUMENT TITLE: ${docContent.title}

FULL DOCUMENT CONTENT (analyze ALL of this):
${docText}
${rulesContext}

Analyze the ENTIRE document above and find ALL issues. Be thorough.`;

  const messages = [
    { role: 'user', parts: [{ text: systemPrompt }] },
    { role: 'user', parts: [{ text: userPrompt }] }
  ];

  const response = callGemini(messages);
  
  let parsedIssues = { issues: [], summary: '' };
  try {
    let clean = response.trim().replace(/```json\s*/g, '').replace(/```\s*/g, '');
    const match = clean.match(/\{[\s\S]*\}/);
    if (match) {
      parsedIssues = JSON.parse(match[0]);
      Logger.log(`Found ${parsedIssues.issues.length} issues`);
    }
  } catch (e) {
    Logger.log('Parse error: ' + e);
  }

  let report = `**DOCUMENT:** ${docContent.title}\n**ISSUES:** ${parsedIssues.issues.length}\n\n`;
  parsedIssues.issues.forEach((issue, i) => {
    report += `${i + 1}. "${issue.text.substring(0, 35)}..." - ${issue.issue}\n`;
  });

  return { parsedIssues, fullReport: report };
}

// ============================================
// Analyze Long Documents in Chunks
// ============================================

function analyzeInChunks(docContent, allRules, rulesContext) {
  const chunkSize = 10000; // Characters per chunk
  const text = docContent.text;
  const chunks = [];
  
  // Split into chunks with overlap
  for (let i = 0; i < text.length; i += chunkSize - 500) { // 500 char overlap
    chunks.push(text.substring(i, Math.min(i + chunkSize, text.length)));
  }
  
  Logger.log(`Split into ${chunks.length} chunks`);
  
  const allIssues = [];
  const seenTexts = new Set(); // Avoid duplicates from overlapping chunks
  
  chunks.forEach((chunk, chunkIndex) => {
    Logger.log(`Analyzing chunk ${chunkIndex + 1}/${chunks.length}...`);
    
    const systemPrompt = `You are a content reviewer. Find ALL issues in this text section.

Return ONLY valid JSON:
{
  "issues": [
    {
      "text": "exact text (15-50 chars)",
      "issue": "what's wrong",
      "suggestion": "fix",
      "rule": "rule name"
    }
  ]
}`;

    const userPrompt = `TEXT SECTION ${chunkIndex + 1}:
${chunk}
${rulesContext}

Find ALL issues in this section.`;

    const messages = [
      { role: 'user', parts: [{ text: systemPrompt }] },
      { role: 'user', parts: [{ text: userPrompt }] }
    ];

    try {
      const response = callGemini(messages);
      let clean = response.trim().replace(/```json\s*/g, '').replace(/```\s*/g, '');
      const match = clean.match(/\{[\s\S]*\}/);
      
      if (match) {
        const result = JSON.parse(match[0]);
        
        result.issues.forEach(issue => {
          const textKey = issue.text.trim();
          if (!seenTexts.has(textKey)) {
            seenTexts.add(textKey);
            allIssues.push(issue);
          }
        });
      }
      
      Utilities.sleep(1000); // Rate limiting between chunks
      
    } catch (e) {
      Logger.log(`Error in chunk ${chunkIndex + 1}: ${e.message}`);
    }
  });
  
  Logger.log(`Total issues found across all chunks: ${allIssues.length}`);
  
  const report = `**DOCUMENT:** ${docContent.title}\n**ISSUES:** ${allIssues.length}\n\n`;
  
  return {
    parsedIssues: {
      issues: allIssues,
      summary: `Found ${allIssues.length} issues across the entire document.`
    },
    fullReport: report
  };
}

// ============================================
// Create Native Comment
// ============================================

function createNativeCommentFormatted(fileId, commentText, startIndex) {
  try {
    const url = `https://www.googleapis.com/drive/v3/files/${fileId}/comments?fields=id,content,anchor`;
    
    const payload = {
      content: commentText,
      anchor: JSON.stringify({
        r: `kix.${startIndex}`
      })
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200 || response.getResponseCode() === 201;
    
  } catch (e) {
    return false;
  }
}

// ============================================
// Add Comments (Native + Backup) with Reference Numbers
// ============================================

function addCommentsWithBackup(docId, doc, analysis) {
  try {
    Logger.log('=== ADDING COMMENTS ===\n');
    
    const body = doc.getBody();
    const fullText = body.getText();
    const issues = analysis.parsedIssues.issues || [];
    
    let nativeCommentsAdded = 0;
    let highlighted = 0;
    const processedIssues = [];
    
    // Process each issue
    issues.forEach((issue, index) => {
      try {
        let searchText = issue.text.trim().replace(/\*\*|❌|✅|AVOID:/g, '').trim();
        
        Logger.log(`[${index + 1}/${issues.length}] "${searchText.substring(0, 40)}..."`);
        
        // Find text
        const searchResult = body.findText(searchText);
        
        if (!searchResult) {
          Logger.log(`   ✗ Not found`);
          return;
        }
        
        const element = searchResult.getElement();
        const startOffset = searchResult.getStartOffset();
        const endOffset = searchResult.getEndOffsetInclusive();
        
        // Highlight in yellow
        element.asText().setBackgroundColor(startOffset, endOffset, '#FFF9C4');
        element.asText().setBold(startOffset, endOffset, false);
        
        // Add reference number RIGHT AFTER the highlighted text
        const refNumber = `[${index + 1}]`;
        try {
          const insertPos = endOffset + 1;
          element.asText().insertText(insertPos, refNumber);
          
          // Format the reference number
          const refEnd = insertPos + refNumber.length - 1;
          element.asText().setFontSize(insertPos, refEnd, 9);
          element.asText().setForegroundColor(insertPos, refEnd, '#D32F2F');
          element.asText().setBold(insertPos, refEnd, false);
          element.asText().setBackgroundColor(insertPos, refEnd, null); // Remove yellow background from number
        } catch (e) {
          Logger.log(`   Could not add reference number: ${e.message}`);
        }
        
        highlighted++;
        
        // Calculate position for native comment
        const elementText = element.asText().getText();
        const elementStartInDoc = fullText.indexOf(elementText);
        const absoluteStart = elementStartInDoc + startOffset + 1;
        
        // Format comment with reference number
        const commentText = `💬 Comment [${index + 1}]

📄 Referenced Text:
"${searchText}"

📝 Issue:
${issue.issue}

💡 Suggestion:
${issue.suggestion}

📋 Rule Reference:
${issue.rule}`;
        
        // Try native comment
        const success = createNativeCommentFormatted(docId, commentText, absoluteStart);
        
        if (success) {
          nativeCommentsAdded++;
          Logger.log(`   ✓ Native comment [${index + 1}] added!`);
        } else {
          Logger.log(`   ⚠ Native comment failed`);
        }
        
        // Store for backup section
        processedIssues.push({
          index: index + 1,
          text: searchText,
          issue: issue.issue,
          suggestion: issue.suggestion,
          rule: issue.rule
        });
        
        Utilities.sleep(500);
        
      } catch (e) {
        Logger.log(`   ✗ Error: ${e.message}`);
      }
    });
    
    Logger.log(`\n=== RESULTS ===`);
    Logger.log(`Highlighted: ${highlighted}`);
    Logger.log(`Native comments: ${nativeCommentsAdded}`);
    
    // Add backup section at end if needed
    if (processedIssues.length > 0 && nativeCommentsAdded < processedIssues.length) {
      Logger.log('Adding backup section at end...');
      addBackupSection(doc, processedIssues, analysis.parsedIssues.summary);
    }
    
    return highlighted;
    
  } catch (error) {
    Logger.log(`Error: ${error}`);
    return 0;
  }
}

// ============================================
// Add Backup Section at End
// ============================================

function addBackupSection(doc, issues, summary) {
  try {
    const body = doc.getBody();
    
    body.appendPageBreak();
    
    const sep = body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    sep.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const header = body.appendParagraph('AI REVIEW COMMENTS');
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    header.setForegroundColor('#1976D2');
    
    const time = body.appendParagraph(new Date().toLocaleString());
    time.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    time.setFontSize(10);
    time.setForegroundColor('#666666');
    
    body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    
    issues.forEach((item) => {
      const ref = body.appendParagraph(`💬 Comment [${item.index}]`);
      ref.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      ref.setForegroundColor('#D32F2F');
      
      body.appendParagraph('');
      body.appendParagraph('📄 Referenced Text:').setBold(true);
      
      const quote = body.appendParagraph(`"${item.text}"`);
      quote.setItalic(true);
      quote.setIndentStart(30);
      quote.setForegroundColor('#666666');
      
      body.appendParagraph('');
      body.appendParagraph('📝 Issue:').setBold(true).setForegroundColor('#E53935');
      body.appendParagraph(item.issue).setIndentStart(30);
      
      body.appendParagraph('');
      body.appendParagraph('💡 Suggestion:').setBold(true).setForegroundColor('#1976D2');
      body.appendParagraph(item.suggestion).setIndentStart(30).setForegroundColor('#1976D2');
      
      body.appendParagraph('');
      body.appendParagraph('📋 Rule Reference:').setBold(true).setFontSize(10);
      body.appendParagraph(item.rule).setIndentStart(30).setFontSize(10).setItalic(true);
      
      body.appendParagraph('');
      body.appendParagraph('─────────────────────────────────').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#CCC');
      body.appendParagraph('');
    });
    
    if (summary) {
      body.appendParagraph('📊 SUMMARY').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(summary);
    }
    
    Logger.log('✓ Backup section added');
    
  } catch (e) {
    Logger.log(`Error adding backup: ${e.message}`);
  }
}

// ============================================
// Main Review Function
// ============================================

function reviewDocumentWithHighlighting(docId, conversationHistory = []) {
  try {
    Logger.log('=== DOCUMENT REVIEW ===\n');
    
    const doc = DocumentApp.openById(docId);
    const docContent = {
      success: true,
      title: doc.getName(),
      text: doc.getBody().getText()
    };
    
    Logger.log(`✓ Document: ${docContent.title}`);

    const pineconeRules = retrieveFromPinecone(docContent.text, 10);
    const driveRules = searchRulesInDrive(docContent.text);
    
    Logger.log(`✓ Rules: ${pineconeRules.length + driveRules.length}`);

    const analysis = analyzeForHighlighting(docContent, { 
      pinecone: pineconeRules, 
      drive: driveRules 
    });
    
    Logger.log(`✓ Issues: ${analysis.parsedIssues.issues.length}\n`);

    const commentsAdded = addCommentsWithBackup(docId, doc, analysis);

    return {
      success: true,
      message: analysis.fullReport,
      documentTitle: docContent.title,
      commentsAdded: commentsAdded
    };

  } catch (error) {
    Logger.log(`ERROR: ${error}`);
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}

// ============================================
// Update Main Process Function
// ============================================

function processMessageWithComments(userMessage, conversationHistory = [], addComments = false) {
  try {
    const isDocReview = isDocumentReviewRequest(userMessage);
    
    if (isDocReview) {
      const docId = extractDocumentId(userMessage);
      if (docId) {
        if (addComments) {
          return reviewDocumentWithHighlighting(docId, conversationHistory);
        } else {
          return reviewDocument(docId, conversationHistory);
        }
      }
    }
    
    return handleConversationalQuery(userMessage, conversationHistory);
    
  } catch (error) {
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}

// ============================================
// Test Function
// ============================================

function testHighlighting() {
  const docUrl = 'https://docs.google.com/document/d/1cysldEwyMHqfSJOMdNKEm26tJPDFvJhTgb6BNA5HYN0/edit';
  const docId = extractDocumentId(docUrl);
  
  const result = reviewDocumentWithHighlighting(docId, []);
  
  Logger.log('\n=== RESULT ===');
  Logger.log(`Success: ${result.success}`);
  Logger.log(`Issues highlighted: ${result.commentsAdded}`);
  
  return result;
}

// ============================================
// HANDLER: Analyze Document With Comments
// ============================================

function handleAnalyzeWithComments(docId, conversationHistory) {
  try {
    Logger.log('🔍 Starting document analysis with comments for: ' + docId);
    
    // Check API configuration
    if (!CHAT_CONFIG.GEMINI_API_KEY) {
      return {
        success: false,
        message: '❌ GEMINI_API_KEY not configured. Please set it in Script Properties.'
      };
    }
    
    // Inform user that analysis is starting
    const startMessage = '🔄 Starting comprehensive document analysis...\n\n' +
                        '**This may take 1-2 minutes. I\'m:**\n' +
                        '• Reading your document\n' +
                        '• Checking against writing rules\n' +
                        '• Identifying issues\n' +
                        '• Adding highlights and comments\n\n' +
                        'Please wait...';
    
    // Call the analysis function from your existing code
    const result = reviewDocumentWithHighlighting(docId, conversationHistory);
    
    if (result.success) {
      const enhancedMessage = `✅ **Analysis Complete!**\n\n` +
                             `📄 **Document:** ${result.documentTitle}\n` +
                             `🎯 **Issues Found:** ${result.commentsAdded}\n\n` +
                             `**What I did:**\n` +
                             `• Highlighted ${result.commentsAdded} issues in yellow\n` +
                             `• Added numbered reference markers [1], [2], etc.\n` +
                             `• Created native Google Doc comments\n` +
                             `• Added a detailed backup section at the end\n\n` +
                             `**Next Steps:**\n` +
                             `1. Open your document to see the highlights\n` +
                             `2. Click on the [1], [2] markers to see comments\n` +
                             `3. Review the backup section at the end\n` +
                             `4. Make edits based on the suggestions\n\n` +
                             `Want me to analyze another document or answer questions about the issues found?`;
      
      return {
        success: true,
        message: enhancedMessage,
        documentTitle: result.documentTitle,
        issuesFound: result.commentsAdded
      };
    } else {
      return {
        success: false,
        message: `❌ Analysis failed: ${result.message}\n\n` +
                `**Common issues:**\n` +
                `• Document not accessible (check sharing settings)\n` +
                `• Invalid document ID\n` +
                `• API configuration problems\n\n` +
                `Need help troubleshooting?`
      };
    }
    
  } catch (error) {
    Logger.log('Error in handleAnalyzeWithComments: ' + error.toString());
    return {
      success: false,
      message: `❌ Error during analysis: ${error.message}\n\n` +
              `This might be due to:\n` +
              `• Permission issues with the document\n` +
              `• API rate limits\n` +
              `• Configuration problems\n\n` +
              `Would you like me to help you troubleshoot?`
    };
  }
}

// ============================================
// HANDLER: Regular Document Review (No Comments)
// ============================================

function handleDocumentReview(docId, conversationHistory) {
  try {
    Logger.log('📄 Starting document review (no comments) for: ' + docId);
    
    // Use the existing reviewDocument function
    const result = reviewDocument(docId, conversationHistory);
    
    if (result.success) {
      const enhancedMessage = `✅ **Document Review Complete**\n\n` +
                             `📄 **Document:** ${result.documentTitle}\n\n` +
                             `**Analysis:**\n${result.message}\n\n` +
                             `**Want me to add comments to the document?**\n` +
                             `Say: "Analyze with comments https://docs.google.com/document/d/${docId}"`;
      
      return {
        success: true,
        message: enhancedMessage
      };
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error in handleDocumentReview: ' + error.toString());
    return {
      success: false,
      message: `❌ Error reviewing document: ${error.message}`
    };
  }
}

function detectUserIntent(message) {
  const lower = message.toLowerCase();
  
  // Extract new documents
  if (lower.match(/extract|scan|check.*new|update.*docs|newly added|get.*new.*documents/)) {
    return { type: 'extract_new_docs' };
  }
  
  // ANALYZE DOCUMENT WITH COMMENTS (NEW)
  if (lower.match(/analyze.*comment|review.*comment|check.*comment|add.*comment|highlight.*issue/)) {
    const docMatch = message.match(/https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/);
    if (docMatch) {
      return { 
        type: 'analyze_with_comments',
        docId: docMatch[1]
      };
    }
    return { type: 'analyze_with_comments_prompt' };
  }
  
  // Regular document review (without comments)
  if (lower.match(/review|analyze|check|feedback|docs\.google\.com/)) {
    const docMatch = message.match(/https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/);
    if (docMatch) {
      return { 
        type: 'review_document',
        docId: docMatch[1]
      };
    }
  }
  
  // Sync to Pinecone
  if (lower.match(/sync|upload.*pinecone|index|embed/)) {
    return { type: 'sync_to_pinecone' };
  }
  
  // Query specific document
  const docNameMatch = lower.match(/(?:document|doc|file).*["']([^"']+)["']|about\s+(.+?)\s+document/);
  if (docNameMatch) {
    return { 
      type: 'query_specific_doc',
      docName: docNameMatch[1] || docNameMatch[2]
    };
  }
  
  // Tab analysis
  if (lower.match(/tabs?|sections?|structure|outline/)) {
    return { type: 'analyze_tabs' };
  }
  
  // General document query
  if (lower.match(/how many|list|show|find|search|documents?|comments?|revisions?/)) {
    return { type: 'query_documents' };
  }
  
  return { type: 'general_chat' };
}

function processChatMessage(userMessage, conversationHistory = []) {
  try {
    Logger.log('Processing message: ' + userMessage);
    
    // Detect intent
    const intent = detectUserIntent(userMessage);
    Logger.log('Detected intent: ' + intent.type);
    
    let result;
    
    switch(intent.type) {
      case 'extract_new_docs':
        result = handleExtractNewDocs();
        break;
      
      // NEW: Document analysis with comments
      case 'analyze_with_comments':
        if (intent.docId) {
          result = handleAnalyzeWithComments(intent.docId, conversationHistory);
        } else {
          result = {
            success: true,
            message: '📝 Please provide the Google Doc URL you want me to analyze and add comments to.\n\nExample: "Analyze with comments https://docs.google.com/document/d/YOUR_DOC_ID"'
          };
        }
        break;
      
      case 'analyze_with_comments_prompt':
        result = {
          success: true,
          message: '📝 I can analyze documents and add inline comments!\n\n**How to use:**\nJust share the Google Doc URL and say:\n• "Analyze this doc with comments: [URL]"\n• "Review and highlight issues: [URL]"\n• "Add comments to this document: [URL]"\n\nI\'ll:\n✅ Find all issues against your writing rules\n✅ Highlight problems in yellow\n✅ Add numbered reference markers\n✅ Create native Google Doc comments\n✅ Add a backup summary at the end\n\nWhat document would you like me to analyze?'
        };
        break;
      
      // Regular review without comments
      case 'review_document':
        if (intent.docId) {
          result = handleDocumentReview(intent.docId, conversationHistory);
        } else {
          result = {
            success: true,
            message: '📄 Please share the Google Doc URL you want me to review.'
          };
        }
        break;
      
      case 'query_documents':
        result = handleDocumentQuery(userMessage, intent, conversationHistory);
        break;
      
      case 'query_specific_doc':
        result = handleSpecificDocQuery(userMessage, intent, conversationHistory);
        break;
      
      case 'analyze_tabs':
        result = handleTabAnalysis(userMessage, intent, conversationHistory);
        break;
      
      case 'sync_to_pinecone':
        result = handleSyncToPinecone();
        break;
      
      case 'general_chat':
      default:
        result = handleGeneralQuery(userMessage, conversationHistory);
        break;
    }
    
    // Ensure result has the required structure
    if (!result || typeof result !== 'object') {
      return {
        success: false,
        message: 'Invalid response from handler'
      };
    }
    
    // Ensure success field exists
    if (result.success === undefined) {
      result.success = true;
    }
    
    // Ensure message field exists
    if (!result.message) {
      result.message = 'No response generated';
    }
    
    Logger.log('Returning result');
    return result;
    
  } catch (error) {
    Logger.log('Error in processChatMessage: ' + error.toString());
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}
