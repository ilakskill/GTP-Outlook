Office.onReady(() => {
  document.addEventListener("DOMContentLoaded", function () {
    const keyInputDiv = document.getElementById("key-input");
    const resultsDiv = document.getElementById("results");
    const apiKeyInput = document.getElementById("apiKeyInput");
    const saveKeyButton = document.getElementById("saveKeyButton");

    const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
    const OPENAI_MODEL = "gpt-4";
    const SYSTEM_MESSAGE =
      "You are a project management assistant. Respond only with JSON.";
    const COOKIE_NAME = "openai_api_key";
    const COOKIE_EXPIRY_DAYS = 30;

      // Helper function to set a cookie
    function setCookie(name, value, days) {
        console.log("Before Setting Cookie: ", document.cookie)
        const expires = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toUTCString();
         const cookieString = `${name}=${value}; expires=${expires}; path=/; Secure; SameSite=None`;
         console.log("Setting Cookie String:", cookieString);
         document.cookie = cookieString;
         console.log("After setting cookie:", document.cookie);
    }


    // Helper function to get a cookie
    function getCookie(name) {
        console.log("Before Getting Cookie: ", document.cookie)
        const match = document.cookie.match(new RegExp(`(^| )${name}=([^;]+)`));
        const cookieValue =  match ? match[2] : null;
        console.log(`Retrieved Cookie: ${name}, Value: ${cookieValue}`)
        return cookieValue;
    }


    function setLoadingState(isLoading) {
      resultsDiv.innerHTML = isLoading ? "<p>Loading...</p>" : "";
    }

    function displayError(message) {
      resultsDiv.innerHTML = `<p style="color: red;">Error: ${message}</p>`;
      console.error("Error: ", message);
    }

    function validateApiKey(apiKey) {
        const regex = /^sk-[a-zA-Z0-9]{48}$/;
        const result = regex.test(apiKey);
        console.log("Validating API key: ", result)
        return result;
    }

    async function fetchEmailBody() {
      return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync("text", (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(`Error retrieving email body: ${result.error.message}`);
          }
        });
      });
    }

      async function analyzeEmail(apiKey, subject, body) {
          const payload = {
              model: OPENAI_MODEL,
              messages: [
                  { role: "system", content: SYSTEM_MESSAGE },
                  {
                      role: "user",
                      content: `Analyze the following email to determine the project and site mentioned. Respond only with JSON in this format: {"project": "project_name", "site": "site_name"}. \n\nSubject: ${subject}\n\nBody: ${body}`
                  }
              ]
          };

        try {
          const response = await fetch(OPENAI_API_URL, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${apiKey}`,
            },
            body: JSON.stringify(payload),
          });

          if (!response.ok) {
            const errorData = await response.json();
            const errorMsg = errorData.error ? errorData.error.message : `OpenAI API Error: ${response.status} ${response.statusText}`;
            throw new Error(errorMsg);
          }

          const data = await response.json();
          const rawContent = data.choices[0].message.content;

          try {
            return JSON.parse(rawContent);
          } catch (error) {
            console.error("Failed to parse OpenAI response as JSON:", rawContent, error);
            throw new Error(
              "Invalid JSON response from OpenAI. Ensure the model is responding in the correct format."
            );
          }

        } catch (error) {
            console.error("OpenAI API Error:", error);
            throw error;
        }
    }

      async function startEmailAnalysis(apiKey) {
        setLoadingState(true);

        if (!Office.context || !Office.context.mailbox) {
            displayError("This add-in can only be used in Outlook.");
              return;
          }

          if(!Office.context.mailbox.item){
             displayError("No email selected.");
              return;
          }

        try {
            const body = await fetchEmailBody();
              if(!body){
                 displayError("Email body is empty.");
                   return;
              }
            const analysisResults = await analyzeEmail(
            apiKey,
            Office.context.mailbox.item.subject,
            body
            );
            resultsDiv.innerHTML = `
                <h2>Analysis Results</h2>
                <p><strong>Project:</strong> ${analysisResults.project}</p>
                <p><strong>Site:</strong> ${analysisResults.site}</p>
            `;
        } catch (error) {
            displayError(error.message);
        } finally {
            setLoadingState(false);
        }
      }

      // Check for stored API key
       console.log("before cookie check on startup: ", document.cookie)
      const storedApiKey = getCookie(COOKIE_NAME);
      if (storedApiKey) {
        keyInputDiv.style.display = "none";
        resultsDiv.style.display = "block";
          startEmailAnalysis(storedApiKey);
      }

      // Save API key and start analysis
      saveKeyButton.addEventListener("click", async () => {
        const apiKey = apiKeyInput.value.trim();

        if (!apiKey) {
          displayError("Please enter an API key.");
          return;
        }

          if(!validateApiKey(apiKey)){
            displayError("Please enter a valid OpenAI API Key (sk-...).");
            return;
        }
        setCookie(COOKIE_NAME, apiKey, COOKIE_EXPIRY_DAYS);
        keyInputDiv.style.display = "none";
        resultsDiv.style.display = "block";
          await startEmailAnalysis(apiKey);
      });
  });
});
