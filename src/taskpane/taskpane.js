/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { analyzeEmailWithOpenAI, getApiKey, storeApiKey } from '../services/openai-service.js';
import { formatDateWithWeekday, formatDateRangeWithWeekdays, isValidApiKeyFormat, storeInRoamingSettings, getFromRoamingSettings } from '../helpers/utilities.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("analyze-button").onclick = analyzeEmail;
  }
});

/**
 * Analyzes the current email using OpenAI API
 */
async function analyzeEmail() {
  try {
    // Show loading spinner
    document.querySelector(".ms-progress").style.display = "block";
    document.getElementById("analysis-result").style.display = "none";
    document.querySelector(".ms-MessageBar").style.display = "none";

    // Get the current email item
    const item = Office.context.mailbox.item;
    
    // Get email body
    const body = await getEmailBody(item);
    
    // Get email headers (subject, from, to, cc)
    const headers = await getEmailHeaders(item);
    
    // Combine headers and body for analysis
    const emailContent = `Subject: ${headers.subject}\nFrom: ${headers.from}\nTo: ${headers.to}\nCC: ${headers.cc || ''}\n\n${body}`;
    
    // Call OpenAI API for analysis
    const analysisResult = await callOpenAIAPI(emailContent);
    
    // Display the result
    displayResult(analysisResult);
    
    // Insert the analysis at the top of the email
    await insertAnalysisInEmail(analysisResult);
  } catch (error) {
    showError(error.message);
  } finally {
    // Hide loading spinner
    document.querySelector(".ms-progress").style.display = "none";
  }
}

/**
 * Gets the email body
 * @param {Office.MessageRead} item - The current email item
 * @returns {Promise<string>} - The email body
 */
function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email body: " + result.error.message));
      }
    });
  });
}

/**
 * Gets the email headers
 * @param {Office.MessageRead} item - The current email item
 * @returns {Promise<object>} - The email headers
 */
function getEmailHeaders(item) {
  try {
    // Safely extract subject
    const subject = item.subject || "[Kein Betreff]";
    
    // Safely extract from address
    let fromAddress = "";
    if (item.from) {
      try {
        fromAddress = item.from.emailAddress || "";
      } catch (e) {
        console.warn("Error extracting from address:", e);
      }
    }
    
    // Safely extract to addresses
    let toAddresses = "";
    if (item.to && Array.isArray(item.to)) {
      try {
        toAddresses = item.to
          .filter(recipient => recipient && recipient.emailAddress)
          .map(recipient => recipient.emailAddress)
          .join("; ");
      } catch (e) {
        console.warn("Error extracting to addresses:", e);
      }
    }
    
    // Safely extract cc addresses
    let ccAddresses = "";
    if (item.cc && Array.isArray(item.cc)) {
      try {
        ccAddresses = item.cc
          .filter(recipient => recipient && recipient.emailAddress)
          .map(recipient => recipient.emailAddress)
          .join("; ");
      } catch (e) {
        console.warn("Error extracting cc addresses:", e);
      }
    }
    
    return {
      subject: subject,
      from: fromAddress,
      to: toAddresses,
      cc: ccAddresses
    };
  } catch (error) {
    console.error("Error getting email headers:", error);
    // Return default values if there's an error
    return {
      subject: "[Fehler beim Lesen des Betreffs]",
      from: "",
      to: "",
      cc: ""
    };
  }
}

/**
 * Gets the OpenAI API key from storage or prompts the user to enter it
 * @returns {Promise<string>} - The API key
 */
async function getOrPromptForApiKey() {
  try {
    // First try to get the API key from roaming settings (more reliable in OWA)
    let apiKey = getFromRoamingSettings('openai_api_key');
    
    // If not found in roaming settings, try localStorage (might not work in OWA)
    if (!apiKey) {
      try {
        apiKey = getApiKey();
      } catch (error) {
        console.warn("Error accessing localStorage:", error);
        // Continue with null apiKey, will prompt user below
      }
    }
    
    // If no API key is found, prompt the user to enter it
    if (!apiKey) {
      apiKey = prompt("Please enter your OpenAI API key:");
      
      // Handle case where user cancels the prompt
      if (!apiKey) {
        throw new Error("API key is required to analyze emails.");
      }
      
      // Validate the API key format
      if (!isValidApiKeyFormat(apiKey)) {
        throw new Error("Invalid API key format. Please enter a valid OpenAI API key.");
      }
      
      // Try to store the API key in both places
      try {
        storeApiKey(apiKey);
      } catch (error) {
        console.warn("Could not store API key in localStorage:", error);
        // Continue anyway, we'll still try roaming settings
      }
      
      // Store in roaming settings (returns a Promise)
      storeInRoamingSettings('openai_api_key', apiKey)
        .then(success => {
          if (!success) {
            console.warn("Could not store API key in roaming settings");
          }
        })
        .catch(error => {
          console.warn("Error storing API key in roaming settings:", error);
        });
      
      // Even if storage fails, we'll still use the key for this session
    }
    
    return apiKey;
  } catch (error) {
    console.error("Error in getOrPromptForApiKey:", error);
    throw new Error("Failed to get API key: " + error.message);
  }
}

/**
 * Calls the OpenAI API to analyze the email
 * @param {string} emailContent - The email content to analyze
 * @returns {Promise<string>} - The analysis result
 */
async function callOpenAIAPI(emailContent) {
  try {
    // Get the API key
    const apiKey = await getOrPromptForApiKey();
    
    // Call the OpenAI API
    let result = await analyzeEmailWithOpenAI(emailContent, apiKey);
    
    // Process the result to add weekday abbreviations to dates
    // This regex looks for date patterns like DD.MM.YYYY
    const datePattern = /\d{1,2}\.\d{1,2}\.\d{4}/g;
    const dateRangePattern = /\d{1,2}\.\d{1,2}\.\d{4}-\d{1,2}\.\d{1,2}\.\d{4}/g;
    
    // First replace date ranges
    result = result.replace(dateRangePattern, match => formatDateRangeWithWeekdays(match));
    
    // Then replace individual dates (that aren't part of ranges)
    result = result.replace(datePattern, match => formatDateWithWeekday(match));
    
    return result;
  } catch (error) {
    console.error("Error calling OpenAI API:", error);
    throw new Error("Failed to analyze email: " + error.message);
  }
}

/**
 * Displays the analysis result in the taskpane
 * @param {string} result - The analysis result
 */
function displayResult(result) {
  const resultElement = document.getElementById("analysis-result");
  resultElement.textContent = result;
  resultElement.style.display = "block";
}

/**
 * Inserts the analysis at the top of the email
 * @param {string} analysis - The analysis result
 */
async function insertAnalysisInEmail(analysis) {
  // Note: In Outlook Add-ins, modifying the content of an email directly is limited
  // This is a simplified approach that may not work in all scenarios
  // A more robust solution would involve creating a custom property or using other Outlook APIs
  
  try {
    // Get the current item
    const item = Office.context.mailbox.item;
    
    // Create a formatted analysis block with the requested styling
    const formattedAnalysis = `
<div style="font-family: Arial, sans-serif; font-weight: bold; color: #00008B; margin-bottom: 20px; padding: 10px; border: 1px solid #ccc; border-radius: 4px;">
${analysis.replace(/\n/g, '<br>')}
</div>
<hr>
`;
    
    // Set a custom property to indicate that we've analyzed this email
    // This could be used in a production scenario to avoid duplicate analyses
    item.loadCustomPropertiesAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        try {
          const props = result.value;
          props.set("vertretungsanalyse_processed", true);
          props.set("vertretungsanalyse_result", analysis);
          
          // Add error handling for saveAsync
          props.saveAsync(function(saveResult) {
            if (saveResult.status !== Office.AsyncResultStatus.Succeeded) {
              console.error("Error saving custom properties:", saveResult.error.message);
            }
          });
        } catch (error) {
          console.error("Error setting custom properties:", error);
        }
      } else {
        console.error("Error loading custom properties:", result.error.message);
      }
    });
    
    // In a real implementation, you would need to use appropriate Outlook APIs
    // to modify the email content, which might require different approaches
    // depending on the Outlook version and context
    
    console.log("Analysis would be inserted at the top of the email:", analysis);
    
    // Show a success message in the taskpane
    const messageBar = document.querySelector(".ms-MessageBar");
    messageBar.classList.remove("ms-MessageBar--error");
    messageBar.classList.add("ms-MessageBar--success");
    messageBar.querySelector(".ms-MessageBar-text").textContent = 
      "Analyse wurde erfolgreich durchgeführt. In einer vollständigen Implementierung würde diese Analyse am Anfang der E-Mail eingefügt werden.";
    messageBar.style.display = "block";
  } catch (error) {
    showError("Failed to insert analysis in email: " + error.message);
  }
}

/**
 * Shows an error message
 * @param {string} message - The error message
 */
function showError(message) {
  const messageBar = document.querySelector(".ms-MessageBar");
  messageBar.classList.remove("ms-MessageBar--success");
  messageBar.classList.add("ms-MessageBar--error");
  messageBar.querySelector(".ms-MessageBar-text").textContent = message;
  messageBar.style.display = "block";
}
