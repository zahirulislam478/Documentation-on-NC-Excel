# Project Documentation

## Overview

This project implements functionality to generate and download an Excel file using data from an HTML table. The solution consists of JavaScript code on the frontend to gather the data and ASP.NET MVC code on the backend to process the data and generate the Excel file using the EPPlus library.

## JavaScript Code

### Function: `DownloadExcelFile()`

This function collects data from an HTML table, serializes it into JSON, and submits it to the server to generate an Excel file.

#### Steps:

1. **Variable Initialization**:
    - `postArrayUI` is initialized as an empty array to hold the selected data.

2. **Table Row Iteration**:
    - Iterate through each row of the table (`#dataTable`), excluding the header row.
    - Check if the checkbox (`.singleCheck`) in the row is checked.
    - If checked, collect `ContactNo` and `SmsText` from the row and push an object containing these into `postArrayUI`.

3. **Validation**:
    - If `postArrayUI` is empty, display an error message using SweetAlert and exit the function.

4. **Form Creation and Submission**:
    - Create a hidden form with a POST method targeting `/NC/AttendanceNotice/GenerateExcel`.
    - Add the serialized `postArrayUI` as a hidden input field.
    - Append the form to the body, submit it, and then remove it.

#### JavaScript Code:

```javascript
function DownloadExcelFile() {
    var postArrayUI = [];
    // Iterate through each row of the table
    $("#dataTable tr").not(":first").each(function () {
        var isChked = $(this).find(".singleCheck").is(':checked');
        if (isChked) {
            var singleObj = {
                ContactNo: $(this).find("td:eq(7)").text(),
                SmsText: $(this).find("td:eq(5)").text()
            };
            postArrayUI.push(singleObj);
        }
    });

    if (postArrayUI.length === 0) {
        swal('Sorry!!', 'No data found', 'error');
        return false;
    }

    var urlToCall = "/NC/AttendanceNotice/GenerateExcel";
    var form = $('<form method="POST" action="' + urlToCall + '">');
    form.append($('<input type="hidden" name="postArrayUI" value=\'' + JSON.stringify(postArrayUI) + '\'>'));
    $('body').append(form);
    form.submit();
    form.remove();
}
