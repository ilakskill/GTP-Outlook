
            Office.onReady(() => {
    const resultsDiv = document.getElementById("results");

    // Fetch email details
    const item = Office.context.mailbox.item;

    if (item) {
        const subject = item.subject;
        item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emailBody = result.value;

                // Send the email data to OpenAI
                analyzeEmail(subject, emailBody)
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
                console.error("Email Body Retrieval Error:", result.error);
            }
        });
    } else {
        resultsDiv.innerHTML = `<p>No email selected.</p>`;
    }
});

// Function to analyze email using OpenAI
async function analyzeEmail(subject, body) {
    const payload = {
        model: "gpt-4",
        messages: [
            { role: "system", content: "You are a project management assistant." },
            { role: "user", content: `Analyze the following email thread to determine the project and site mentioned. \n\nSubject: ${subject}\n\nBody: ${body}` }
        ]
    };

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `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`
        },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    const analysis = data.choices[0].message.content;

    // Assume OpenAI returns JSON with 'project' and 'site' fields
    return JSON.parse(analysis);
}
