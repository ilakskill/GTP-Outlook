// This script runs in the new window
document.addEventListener("DOMContentLoaded", () => {
    const contentDiv = document.getElementById("content");
    contentDiv.textContent = "Your add-in is working!";

    // Example: Make an API call or process some data here
    // fetch("https://example.com/api")
    //     .then(response => response.json())
    //     .then(data => {
    //         contentDiv.textContent = JSON.stringify(data);
    //     })
    //     .catch(error => {
    //         contentDiv.textContent = "Error loading data.";
    //         console.error(error);
    //     });
});
