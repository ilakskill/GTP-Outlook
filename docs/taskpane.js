Office.onReady(() => {
  document.addEventListener("DOMContentLoaded", function () {

    function setCookieTest(name, value) {
      const expires = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toUTCString();
      document.cookie = `${name}=${value}; expires=${expires}; path=/; Secure; SameSite=None`;
      console.log("Setting Cookie:", document.cookie);
    }

    function getCookieTest(name) {
      console.log("Getting Cookie:", document.cookie);
       const match = document.cookie.match(new RegExp(`(^| )${name}=([^;]+)`));
       const cookieValue =  match ? match[2] : null;
       console.log(`Retrieved Cookie: ${name}, Value: ${cookieValue}`)
        return cookieValue;
     }
    setCookieTest("test_cookie", "test_value");
    getCookieTest("test_cookie");
  });
});
