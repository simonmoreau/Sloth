// Write your Javascript code.

//Diaplay file name before upload
document.getElementById("uploadBtn").onchange = function () {
    document.getElementById("uploadFile").value = this.files[0].name;
};