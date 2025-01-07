document.addEventListener("DOMContentLoaded", function () {
    const contentDiv = document.getElementById("content");

    // Example API call to Smartsheet
    fetch("https://api.smartsheet.com/2.0/sheets/{sheet_id}", {
        method: "GET",
        headers: {
            "Authorization": "Bearer YOUR_API_KEY"
        }
    })
    .then(response => response.json())
    .then(data => {
        // Display results
        contentDiv.innerHTML = `
            <h2>Smartsheet Data</h2>
            <pre>${JSON.stringify(data, null, 2)}</pre>
        `;
    })
    .catch(error => {
        contentDiv.innerHTML = `<p>Error fetching data: ${error.message}</p>`;
        console.error("Error fetching Smartsheet data:", error);
    });
});

