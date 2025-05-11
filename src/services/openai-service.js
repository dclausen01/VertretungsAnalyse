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
      })
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(`OpenAI API error: ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  } catch (error) {
    console.error('Error calling OpenAI API:', error);
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
  
  // This is just a placeholder to show how it might be structured
  localStorage.setItem('openai_api_key', apiKey);
}

/**
 * Retrieves the stored OpenAI API key
 * @returns {string|null} - The stored API key, or null if not found
 */
export function getApiKey() {
  // In a production environment, the API key should be retrieved securely
  // For demonstration purposes, we're using localStorage, but this is not secure
  // In a real implementation, this should be handled by a server-side component
  
  // This is just a placeholder to show how it might be structured
  return localStorage.getItem('openai_api_key');
}
