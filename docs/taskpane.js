                // 6 jan 2025

Office.onReady(() => {
    const resultsDiv = document.getElementById("results");

    // Fetch email details
    const item = Office.context.mailbox.item;

    if (item) {
        const subject = item.subject;
        console.log("Email Subject:", subject);

        item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Email Body Retrieved:", result.value);
                const emailBody = result.value;

                // Send the email data to OpenAI 6 jan 2025
                analyzeEmail(subject, emailBody)
                    .then(response => {
                        console.log("OpenAI Response:", response);
                        resultsDiv.innerHTML = `
                            <h2>Analysis Results</h2>
                            <p><strong>Project:</strong> ${response.project}</p>
                            <p><strong>Site:</strong> ${response.site}</p>
                        `;
                    })
                    .catch(error => {
                        console.error("Error Analyzing Email:", error);
                        resultsDiv.innerHTML = `<p>Error analyzing email: ${error.message}</p>`;
                    });
            } else {
                console.error("Email Body Retrieval Error:", result.error);
                resultsDiv.innerHTML = `<p>Error retrieving email body: ${result.error.message}</p>`;
            }
        });
    } else {
        console.error("No email selected.");
        resultsDiv.innerHTML = `<p>No email selected.</p>`;
    }
});

async function analyzeEmail(subject, body) {
    console.log("Sending Email to OpenAI:", { subject, body });

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
            "Authorization": 'Bearer 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'
        },
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        console.error("OpenAI API Error:", response.status, response.statusText);
        throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    console.log("OpenAI API Raw Response:", data);

    return JSON.parse(data.choices[0].message.content);
}
