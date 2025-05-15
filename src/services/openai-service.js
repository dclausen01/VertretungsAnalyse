/*
 * Service for interacting with the OpenAI API
 */

/**
 * Calls the OpenAI API to analyze an email
 * @param {string} emailContent - The email content to analyze
 * @param {string} apiKey - The OpenAI API key
 * @returns {Promise<string>} - The analysis result
 */
export async function analyzeEmailWithOpenAI(emailContent, apiKey) {
  try {
    // In a production environment, this API call should be made through a secure backend service
    // to protect the API key and handle rate limiting, error handling, etc.
    
    // For demonstration purposes, we're showing how the API call might be structured
    // but in a real implementation, this should be handled by a server-side component
    
    // Check for network connectivity before making the API call
    if (!navigator.onLine) {
      throw new Error("No internet connection. Please check your network and try again.");
    }
    
    // Add timeout to the fetch request to prevent hanging
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000); // 30 second timeout
    
    try {
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: 'gpt-4o-mini', // or another appropriate model
          messages: [
            {
              role: 'system',
              content: `Du bist ein Assistent, der Lehrermails analysiert und strukturierte Regelungsvorschläge für die Schulsoftware UNTIS erstellt.`
            },
            {
              role: 'user',
              content: `Analysiere den folgenden E-Mail-Text und gib eine kompakte Zusammenfassung im folgenden Format aus:

---
FORMAT DER AUSGABE:
<Lehrkraft-Kürzel>: <Typ> <Datum> <ggf. Uhrzeit/Zeitraum> – <Kurzregelung oder Maßnahmen>
---

**Regeln und Hinweise:**
- Verwende das Lehrer:innen-Kürzel nach dem Schema: Erste 4 Buchstaben des Nachnamens + 1. Buchstabe des Vornamens (z. B. Clausen = ClauD), beachte Ausnahmen wie PeMar, PetMa, VosAn.
- Mögliche Typen: *krank*, *Raumbuchung*, *Fortbildung*, *Vertretungsregelung*
- Kürze auf das Wesentliche, formuliere aber klar und verständlich.
- Bei Raumbuchungen: Gib den Raum an und die Startzeit (z. B. „ab 8:15")
- Bei Krankheit/Fortbildung: Gib Zeitraum an und vorhandene Vertretungsregelungen.
- Verwende keine persönlichen Anreden oder E-Mail-Form.
- Ergänze immer auch den Wochentag als Kürzel zu jedem Datum.
- Küchenräume (R115, R118, R015, R018) dürfen nur von SievJ, PeteM, GellR, HantJ genutzt werden – gib eine Warnung aus, wenn das nicht zutrifft.
- Wenn ein Datum oder Kürzel fehlt oder unklar ist, formuliere einen Hinweis wie [Datum fehlt] oder [Kürzel nicht erkennbar].

---

**E-Mail-Text:**
${emailContent}`
            }
          ],
          temperature: 0.7,
          max_tokens: 500
        }),
        signal: controller.signal
      });
      
      clearTimeout(timeoutId); // Clear the timeout if the request completes
      
      if (!response.ok) {
        let errorMessage = `HTTP error ${response.status}: ${response.statusText}`;
        try {
          const errorData = await response.json();
          errorMessage = `OpenAI API error: ${errorData.error?.message || errorMessage}`;
        } catch (jsonError) {
          // If we can't parse the error as JSON, just use the HTTP error
        }
        throw new Error(errorMessage);
      }

      const data = await response.json();
      
      // Validate the response structure
      if (!data.choices || !data.choices[0] || !data.choices[0].message || !data.choices[0].message.content) {
        throw new Error("Invalid response format from OpenAI API");
      }
      
      return data.choices[0].message.content.trim();
    } catch (fetchError) {
      if (fetchError.name === 'AbortError') {
        throw new Error("Request timed out. The server took too long to respond.");
      }
      throw fetchError;
    } finally {
      clearTimeout(timeoutId); // Ensure the timeout is cleared
    }
  } catch (error) {
    console.error('Error calling OpenAI API:', error);
    
    // Provide more user-friendly error messages
    if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
      throw new Error("Network error when connecting to OpenAI. Please check your internet connection and try again.");
    } else if (error.message.includes('401')) {
      throw new Error("Invalid API key. Please check your OpenAI API key and try again.");
    } else if (error.message.includes('429')) {
      throw new Error("OpenAI API rate limit exceeded. Please try again later.");
    } else if (error.message.includes('500')) {
      throw new Error("OpenAI server error. Please try again later.");
    }
    
    throw new Error(`Failed to analyze email: ${error.message}`);
  }
}

/**
 * Stores the OpenAI API key securely
 * @param {string} apiKey - The OpenAI API key to store
 */
export function storeApiKey(apiKey) {
  // In a production environment, the API key should be stored securely
  // For demonstration purposes, we're using localStorage, but this is not secure
  // In a real implementation, this should be handled by a server-side component
  
  try {
    // Check if localStorage is available
    if (typeof localStorage === 'undefined') {
      console.warn('localStorage is not available in this environment');
      return false;
    }
    
    // Test localStorage by setting and getting a test value
    const testKey = '_test_localStorage_';
    try {
      localStorage.setItem(testKey, 'test');
      if (localStorage.getItem(testKey) !== 'test') {
        throw new Error('localStorage test failed');
      }
      localStorage.removeItem(testKey);
    } catch (e) {
      console.warn('localStorage is not working properly:', e);
      return false;
    }
    
    // If we got here, localStorage is available and working
    localStorage.setItem('openai_api_key', apiKey);
    return true;
  } catch (error) {
    console.error('Error storing API key in localStorage:', error);
    return false;
  }
}

/**
 * Retrieves the stored OpenAI API key
 * @returns {string|null} - The stored API key, or null if not found
 */
export function getApiKey() {
  // In a production environment, the API key should be retrieved securely
  // For demonstration purposes, we're using localStorage, but this is not secure
  // In a real implementation, this should be handled by a server-side component
  
  try {
    // Check if localStorage is available
    if (typeof localStorage === 'undefined') {
      console.warn('localStorage is not available in this environment');
      return null;
    }
    
    return localStorage.getItem('openai_api_key');
  } catch (error) {
    console.error('Error retrieving API key from localStorage:', error);
    return null;
  }
}
