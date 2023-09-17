//About section
const aboutButton = document.querySelector(".about");
const aboutDiv = document.querySelector(".about-div");

//Upload form
const fileUploadForm = document.getElementById("upload-form");

//PDF Tools
const changeMailIDPassButton = document.querySelector(".change-sender-id-pass");
const generatePDFButton = document.querySelector(".generate-pdf");
const pdfHelperText = document.querySelector(".pdf-helper-text");
const mailSenderHelperText = document.querySelector(".mail-send-helper-text");
const sendEmailsButton = document.querySelector(".mailer-pdf");

//To set Sender Mail ID and password
const mailTools = document.querySelector(".mail-tools");
const setMailAndPassButton = document.querySelector(".set-mail-pass");
const closeButton = document.querySelector(".close-mail-tools");
const mailHelperText = document.querySelector(".mail-helper-text");

//File Upload Validation
fileUploadForm.addEventListener("submit", function (event) {
  event.preventDefault();
  var formData = new FormData(this);

  fetch("/upload-file", {
    method: "POST",
    body: formData,
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.error) {
        document.getElementById("upload-result").innerHTML =
          "<p>Error: " + data.error + "</p>";
      } else {
        document.getElementById("upload-result").innerHTML =
          "<p> File Uploaded: reports </p>";
      }
    })
    .catch((error) => {
      console.error("Error:", error);
    });
});

function toggleMailTools() {
  mailHelperText.classList.add("invisible");
  document.querySelector("#mail-id").value = "";
  document.querySelector("#mail-pass").value = "";
  mailTools.classList.toggle("invisible");
}

function showAbout() {
  aboutDiv.classList.toggle("invisible");
}

function getPDF() {
  usnStartVar = document.getElementById("usn-from").value.toUpperCase();
  usnEndVar = document.getElementById("usn-to").value.toUpperCase();
  $.ajax({
    url: "/generate-report",
    type: "POST",
    contentType: "application/json",
    data: JSON.stringify({ usnStartVar: usnStartVar, usnEndVar: usnEndVar }),
    success: function (postResponse) {
      if (postResponse.error) {
        pdfHelperText.classList.remove("invisible");
        pdfHelperText.textContent = `Error: ${postResponse.error}, Message: ${postResponse.Message}`;
      } else if (postResponse.operation) {
        pdfHelperText.classList.remove("invisible");
        pdfHelperText.textContent = `Operation: ${postResponse.operation}, Message: ${postResponse.Message}`;
        const zipFilePath = postResponse.zip_file_url;

        // Request the ZIP file
        $.ajax({
          url: "/send-zip",
          type: "POST",
          contentType: "application/json",
          xhrFields: {
            responseType: "blob",
          },
          data: JSON.stringify({ zipFilePath: zipFilePath }),
          success: function (response) {
            // Create a temporary anchor element for download
            const downloadLink = document.createElement("a");
            downloadLink.style.display = "none";
            downloadLink.download = "Reports.zip";
            downloadLink.href = window.URL.createObjectURL(response);
            document.body.appendChild(downloadLink);
            downloadLink.click();

            // Clean up after the download link
            document.body.removeChild(downloadLink);
            window.URL.revokeObjectURL(downloadLink.href);
          },
          error: function (response) {
            console.error(response.error);
          },
        });
      }
    },
    error: function (postError) {
      console.error("Error:", postError);
    },
  });
}

function sendData() {
  const mailID = document.querySelector("#mail-id").value;
  const mailPass = document.querySelector("#mail-pass").value;
  $.ajax({
    url: "/change-mail",
    type: "POST",
    contentType: "application/json",
    //Data to send: mail ID and password
    data: JSON.stringify({ newMail: mailID, newPass: mailPass }),
    success: function (response) {
      //Set Value of p element for success/failure of mail setting
      mailHelperText.classList.remove("invisible");
      if (response.error) {
        mailHelperText.innerHTML = `Error: ${response.error}`;
      }
      if (response.operation) {
        mailHelperText.innerHTML = `Operation: ${response.operation}`;
      }
    },
    error: function (error) {
      mailHelperText.innerHTML = error;
    },
  });
}

function sendEmails() {
  $.ajax({
    url: "/send-mail",
    type: "GET",
    success: function (response) {
      mailSenderHelperText.classList.remove("invisible");
      if (response.error) {
        mailSenderHelperText.innerHTML = `Error: ${response.error}, Message: ${response.Message}`;
      }
      if (response.operation) {
        mailSenderHelperText.innerHTML = `Operation: ${response.operation}, Message: ${response.Message}`;
      }
    },
    error: function (error) {
      mailSenderHelperText.innerHTML = error;
    },
  });
}

aboutButton.addEventListener("click", showAbout);
changeMailIDPassButton.addEventListener("click", toggleMailTools);
generatePDFButton.addEventListener("click", getPDF);
setMailAndPassButton.addEventListener("click", sendData);
closeButton.addEventListener("click", toggleMailTools);
sendEmailsButton.addEventListener("click", sendEmails);
