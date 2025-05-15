/*
 * Simplified version of the taskpane.js file for testing
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("analyze-button").onclick = displaySimpleMessage;
  }
});

/**
 * Displays a simple message in the taskpane
 */
function displaySimpleMessage() {
  try {
    // Show a simple message
    const resultElement = document.getElementById("analysis-result");
    resultElement.textContent = "Vertretungsanalyse Add-in funktioniert! Dies ist eine vereinfachte Version ohne OpenAI API-Integration.";
    resultElement.style.display = "block";
    
    // Show a success message
    const messageBar = document.querySelector(".ms-MessageBar");
    messageBar.classList.remove("ms-MessageBar--error");
    messageBar.classList.add("ms-MessageBar--success");
    messageBar.querySelector(".ms-MessageBar-text").textContent = 
      "Test erfolgreich! Die vereinfachte Version des Add-ins funktioniert.";
    messageBar.style.display = "block";
  } catch (error) {
    // Show an error message
    const messageBar = document.querySelector(".ms-MessageBar");
    messageBar.classList.remove("ms-MessageBar--success");
    messageBar.classList.add("ms-MessageBar--error");
    messageBar.querySelector(".ms-MessageBar-text").textContent = "Fehler: " + error.message;
    messageBar.style.display = "block";
  }
}
