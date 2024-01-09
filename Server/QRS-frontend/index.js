const form = document.querySelector("#form-info");
const loader = document.querySelector("#loader");
const results = document.querySelector("#results");
const overlay = document.querySelector("#overlay");
const dowload = document.querySelector("#save-btn");
const className = document.querySelector("#class-name")

let blobUrl;

// Take over form submission
form.addEventListener("submit", (event) => {
    event.preventDefault();
    sendData();
});

// send form data to server and show loader
async function sendData() {
    const formData = new FormData(form);
  try {
    displayLoader();
    const response = await fetch("http://127.0.0.1:5000/fetch_results", {
      method: "POST",
        // Set the FormData instance as the request body
      body: formData,
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const blob = await response.blob();
    blobUrl = URL.createObjectURL(blob);

    hideLoader();
    displayResult();

  } catch (e) {
    hideLoader();
    alert(e.message);
  }
}

function saveFile(url, filename) {
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "file-name";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// save csv file
dowload.addEventListener("click", () => {
    saveFile(blobUrl, `result-${className.value}.csv`);
    URL.revokeObjectURL(blobUrl);
    hideResult();
    form.reset();
})

function displayLoader(){
    overlay.style.display = "block";
    loader.style.display = "flex";
}

function hideLoader(){
    overlay.style.display = "none";
    loader.style.display = "none";
}

function displayResult(){
    overlay.style.display = "block";
    results.style.display = "flex";
}

function hideResult(){
    overlay.style.display = "none";
    results.style.display = "none";
}

