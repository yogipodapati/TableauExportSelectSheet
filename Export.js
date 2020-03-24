$(document).ready(function() {
    $('#export').click(function() {
        initializeTableau();
    });

});

function initializeTableau() {
    tableau.extensions.initializeAsync().then(function() {
        getSummaryData();
    }, function(err) {
        console.log('Error while Initializing: ' + err.toString());
    });
}

function prepareHeaders(objArr) {
    var header = "";
    objArr.forEach(function(obj) {
        header = header + obj['_fieldName'] + ',';
    });
    return header + "\r\n";
}

function prepareHeadersArr(objArr) {
    //alert(objArr)
    var header = [];
    objArr.forEach(function(obj) {
        header.push(obj['_fieldName']);
    });
    return header;
}

function convertToCSV(objArray) {
    var str = '';
    objArray['_data'].forEach(d1 => {
        let line = '';
        d1.forEach(d2 => {
            if (line != '')
                line = line + ','
            line = line + d2['_formattedValue'];
        });
        str += line + '\r\n';

    });
    return str;
}


function downloadExcel(obj) {
    var csv = convertToCSV(obj);
    var exportedFilenmae = "export.csv";
    var headers = prepareHeaders(obj['_columns']);
    csv = headers + csv;
    var confInfo = "** THIS IS CONFIDENTIAL INFORMATION**";
    csv = csv + '\r\n' + confInfo;
    var blob = new Blob([csv], {
        type: 'text/csv;charset=utf-8;'
    });
    if (navigator.msSaveBlob) { // IE 10+
        navigator.msSaveBlob(blob, exportedFilenmae);
    } else {
        var link = document.createElement("a");
        if (link.download !== undefined) { // feature detection
            // Browsers that support HTML5 download attribute
            var url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", exportedFilenmae);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

}

// This is a handling function that is called anytime a filter is changed in Tableau.
function filterChangedHandler(filterEvent) {
    // Just reconstruct the filters table whenever a filter changes.
    // This could be optimized to add/remove only the different filters.
    getSummaryData();
}

function createSheetNamesDropDown(sheetNames) {
    // there any changes in dependents?
 //   alert()
    var combo = $("#dropdown");
    alert(combo);
    $.each(sheetNames, function(i, el) {

        combo.append(`<option value=${el}> ${el}</option>`);

    });


    // return combo;
    // OR

    //  $("#dropdown").append(combo);
}

function extractSheetNames(dashboard) {
    var sheetnames = [];
    var combo = $("#dropdown");
    dashboard.worksheets.forEach(function(worksheet) {
      combo.append(`<option value=${JSON.stringify(worksheet)}> ${worksheet.name}</option>`);
       // alert("name .... " + worksheet.name)
      //  sheetnames.push(worksheet.name)
    });
    createSheetNamesDropDown(sheetnames);

}

function selectSheet() {
    let selectedSheet = $("#dropdown").val();
    alert(JSON.stringify(selectedSheet))

}

function getSummaryData() {
    //  alert();
    const dashboard = tableau.extensions.dashboardContent.dashboard;

    extractSheetNames(dashboard);
    // createSheetNamesDropDown(sheetnames);
    // Then loop through each worksheet and get its filters, save promise for later.
    dashboard.worksheets.forEach(function(worksheet) {
        //  sheetname -> sheet data
        options = {
            maxRows: 0,
            ignoreAliases: true,
            ignoreSelection: true,
            includeAllColumns: false
        };

        //if(worksheet.name == selectedSheet ){

        worksheet.getSummaryDataAsync(options).then(function(data) {
            // if (worksheet.name === 'AAA')
            // {
            downloadExcel(data);

            ////}
        });
        // }
    });
}