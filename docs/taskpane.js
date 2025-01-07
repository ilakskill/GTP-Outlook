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

  // Check for stored API key
console.log("before cookie check on startup: ", document.cookie)
const storedApiKey = getCookie(COOKIE_NAME);
if (storedApiKey) {
  keyInputDiv.style.display = "none";
  resultsDiv.style.display = "block";
   startEmailAnalysis(storedApiKey);
}
