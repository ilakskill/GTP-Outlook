  Office.onReady(() => {
   console.log("Office.onReady has been called");
    document.addEventListener("DOMContentLoaded", function () {
       console.log("DOMContentLoaded has been called");
       const saveKeyButton = document.getElementById("saveKeyButton");
       console.log('Button', saveKeyButton)
       saveKeyButton.addEventListener("click", async () => {
          console.log('button was clicked');
       });
    });
  });
