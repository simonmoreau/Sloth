// Write your Javascript code.

//Display file name before upload
document.getElementById("uploadBtn").onchange = function () {
    document.getElementById("uploadFile").placeholder = this.files[0].name;
};

document.getElementById('Submit').onclick = function () {
    var uploadBtn = document.getElementById("uploadBtn");
    var validFilesTypes = ["bcf", "bcfzip"];

    if (uploadBtn.files == null || uploadBtn.files.length == 0) {
        document.getElementById("uploadFile").placeholder = "Please select a BFC file";
    }
    else if (!validFilesTypes.includes(uploadBtn.files[0].name.split('.').pop())) {
        document.getElementById("uploadFile").placeholder = "Please select a .bcf or a .bcfzip file";
    }
    else {
        this.form.submit();
    }
}

document.getElementById("WordExportBnt").onclick = function () {

    if (modelcount == 0) {
        var snackbarContainer = document.querySelector('#export-snackbar');

        var data = { timeout: 2000, message: 'Please upload a BCF file' };
        snackbarContainer.MaterialSnackbar.showSnackbar(data);
    }
    else {
        this.form.submit();
    }
};

document.getElementById("ExcelExportBnt").onclick = function () {
    if (modelcount == 0) {
        var snackbarContainer = document.querySelector('#export-snackbar');

        var data = { timeout: 2000, message: 'Please upload a BCF file' };
        snackbarContainer.MaterialSnackbar.showSnackbar(data);
    }
    else {
        this.form.submit();
    }
};
