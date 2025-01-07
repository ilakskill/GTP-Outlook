Office.onReady(() => {
    document.addEventListener("DOMContentLoaded", function () {
        const keyInputDiv = document.getElementById("key-input");
        const resultsDiv = document.getElementById("results");
        const apiKeyInput = document.getElementById("apiKeyInput");
        const saveKeyButton = document.getElementById("saveKeyButton");

        // Helper function to set a cookie
        function setCookie(name, value, days) {
            const expires = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toUTCString();
            const cookieString = `${name}=${value}; expires=${expires}; path=/; Secure; SameSite=None`;
            console.log("Setting Cookie:", cookieString);
            document.cookie = cookieString;
        }

        // Helper function to get a cookie
        function getCookie(name) {
            console.log("Getting Cookies:", document.cookie);
            const match = document.cookie.match(new RegExp(`(^| )${name}=([^;]+)`));
            return match ? match[2] : null;
        }

        // Check for stored API key
        const storedApiKey = getCookie("openai_api_key");
        if (storedApiKey) {
            keyInputDiv.style.display = "none";
            resultsDiv.style.display = "block";
            startEmailAnalysis(storedApiKey);
        }

        // Save API key and start analysis
        saveKeyButton.addEventListener("click", () => {
            console.log("Save Key button clicked.");
            const apiKey = apiKeyInput.value.trim();
            if (apiKey) {
                console.log("API Key entered:", apiKey);
                setCookie("openai_api_key", apiKey, 30); // Save key for 30 days
                keyInputDiv.style.display = "none";
                resultsDiv.style.display = "block";
                startEmailAnalysis(apiKey);
            } else {
                alert("Please enter a valid API key.");
            }
        });

        // Function to analyze the email
        function startEmailAnalysis(apiKey) {
            console.log("Starting email analysis with API Key:", apiKey);
            if (!Office.context || !Office.context.mailbox) {
                console.error("Office.context.mailbox is not available. Ensure the add-in is running in Outlook.");
                resultsDiv.innerHTML = `<p>Error: This add-in can only be used in Outlook.</p>`;
                return;
            }

            const item = Office.context.mailbox.item;
            if (!item) {
                console.error("No email item available.");
                resultsDiv.innerHTML = `<p>Error: No email selected.</p>`;
                return;
            }

            console.log("Analyzing email:", item.subject);
            item.body.getAsync("text", (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Email Body Retrieved:", result.value);
                    const emailBody = result.value;

                    analyzeEmail(apiKey, item.subject, emailBody)
                        .then(response => {
                            console.log("OpenAI Response:", response);
                            resultsDiv.innerHTML = `
                                <h2>Analysis Results</h2>
                                <p><strong>Project:</strong> ${response.project}</p>
                                <p><strong>Site:</strong> ${response.site}</p>
                            `;
                        })
                        .catch(error => {
                            resultsDiv.innerHTML = `<p>Error analyzing email: ${error.message}</p>`;
                            console.error("OpenAI Analysis Error:", error);
                        });
                } else {
                    console.error("Error retrieving email body:", result.error);
                    resultsDiv.innerHTML = `<p>Error retrieving email body: ${result.error.message}</p>`;
                }
            });
        }

        // Function to call OpenAI API
        async function analyzeEmail(apiKey, subject, body) {
            const payload = {
                model: "gpt-4",
                messages: [
                    { role: "system", content: "You are a project management assistant. Respond only with JSON." },
                    {
                        role: "user",
                        content: `Analyze the following email to determine the project and site mentioned. Respond only with JSON in this format: {"project": "project_name", "site": "site_name"}. \n\nSubject: ${subject}\n\nBody: ${body}`
                    }
                ]
            };

            const response = await fetch("https://api.openai.com/v1/chat/completions", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${apiKey}`
                },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                throw new Error(`OpenAI API error: ${response.statusText}`);
            }

            const data = await response.json();
            const rawContent = data.choices[0].message.content;

            try {
                // Attempt to parse the content as JSON
                return JSON.parse(rawContent);
            } catch (error) {
                console.error("Failed to parse OpenAI response as JSON:", rawContent);
                throw new Error("Invalid JSON response from OpenAI. Ensure the model is responding in the correct format.");
            }
        }
    });
});
