function uploadFile() {
  const fileInput = document.getElementById("fileInput");
  const file = fileInput.files[0];

  if (!file) {
    alert("Please select a file.");
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  fetch("/upload", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (response.ok) {
        return response.json();
      }
      throw new Error("File upload failed.");
    })
    .then((data) => {
      alert(data.message);
    })
    .catch((error) => {
      console.error("Error:", error);
      alert("Error uploading file.");
    });
}
