document.addEventListener("DOMContentLoaded", function () {
    const keyInputDiv = document.getElementById("key-input");
    const resultsDiv = document.getElementById("results");
    const apiKeyInput = document.getElementById("apiKeyInput");
    const saveKeyButton = document.getElementById("saveKeyButton");

    // Helper function to set a cookie
    function setCookie(name, value, days) {
        const expires = new Date(Date.now() + days * 24 * 60 * 60 * 1000).toUTCString();
        document.cookie = `${name}=${value}; expires=${expires}; path=/`;
    }

    // Helper function to get a cookie
    function getCookie(name) {
        const match = document.cookie.match(new RegExp(`(^| )${name}=([^;]+)`));
        return match ? match[2] : null;
    }

    // Check for stored API key
    const storedApiKey = getCookie("openai_api_key");
    if (storedApiKey) {
        keyInputDiv.style.display = "none";
        resultsDiv.style.display = "block";
        startEmailAnalysis(storedApiKey); // Start analyzing email
    }

    // Save API key and start analysis
    saveKeyButton.addEventListener("click", () => {
        const apiKey = apiKeyInput.value.trim();
        if (apiKey) {
            setCookie("openai_api_key", apiKey, 30); // Save key for 30 days
            keyInputDiv.style.display = "none";
            resultsDiv.style.display = "block";
            startEmailAnalysis(apiKey); // Start analyzing email
        } else {
            alert("Please enter a valid API key.");
        }
    });

    // Function to analyze the email
    function startEmailAnalysis(apiKey) {
        const item = Office.context.mailbox.item;

        if (item) {
            const subject = item.subject;
            item.body.getAsync("text", (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailBody = result.value;

                    // Send data to OpenAI
                    analyzeEmail(apiKey, subject, emailBody)
                        .then(response => {
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
                    resultsDiv.innerHTML = `<p>Error retrieving email body: ${result.error.message}</p>`;
                }
            });
        } else {
            resultsDiv.innerHTML = `<p>No email selected.</p>`;
        }
    }

    // Function to call OpenAI API
    async function analyzeEmail(apiKey, subject, body) {
        const payload = {
            model: "gpt-4",
            messages: [
                { role: "system", content: "You are a project management assistant." },
                { role: "user", content: `Analyze the following email to determine the project and site mentioned:\n\nSubject: ${subject}\n\nBody: ${body}` }
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
        return JSON.parse(data.choices[0].message.content);
    }
});
