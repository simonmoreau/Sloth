// Write your Javascript code.

//Display file name in the upload form
document.getElementById("uploadBtn").onchange = function () {
    document.getElementById("uploadFile").value = this.files[0].name;
};