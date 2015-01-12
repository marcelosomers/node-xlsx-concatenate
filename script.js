var fs = require("fs")
    xlsx = require("xlsx");

/**
 * Configuration
 * Set Input Directory and Export Directory
 */
var inputDirectory = "input/";
var exportDirectory = "export/";


/**
 * Don't touch anything below!
 */
var csvContent = "";
var output = "";

// Set up our file name
var today = new Date();
var yyyy = today.getFullYear();
var mm = appendLeadingZero(today.getMonth()+1);    // January is 0
var dd = appendLeadingZero(today.getDate());
var hh = appendLeadingZero(today.getHours());
var mm = appendLeadingZero(today.getMinutes());
var ss = appendLeadingZero(today.getSeconds());

function appendLeadingZero(value) {
    if (value < 10) {
        value = "0" + value;
    }
    return value;
}

var exportFileName = yyyy + "-" + mm + "-" + dd + " " + hh + "." + mm + "." + ss + ".csv"


/**
 * The magic starts here!
 */
fs.readdir(inputDirectory, function(err, files) {
    if (err) {
        throw err;
    }

    // Loop over all files in the input directory
    for(var i=0; i < files.length; i++) {

        var fileName = files[i];
        var fileType = fileName.substr(fileName.length - 5);

        // Parse file
        to_csv(file, files[i]);
        
        // Only open xlsx files
        if(fileType === ".xlsx") {
            // Read the file
            var file = xlsx.readFile(inputDirectory + files[i]);

            // parse the file
            parseFile(file, fileName, i);
        }
    }
});


function to_csv(workbook, fileName) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
        if(roa.length > 0){
            result[sheetName] = roa;
        }

        // Add the file name
        fs.appendFile("output.csv", fileName + '\r\n', function(err) {
            if(err) {
                throw err;
            }
        })

        // Add the CSV content
        fs.appendFile("output.csv", roa, function(err) {
            if(err) {
                throw err;
            }
        })
    });

function parseFile(workbook, fileName, loop) {
    var sheet_name_list = workbook.SheetNames;
    var sheet = workbook.Sheets[sheet_name_list[0]];
    var data = xlsx.utils.sheet_to_json(sheet, {header:1});

    // Process the blank cells to use an empty string
    for(var i = 0; i != data.length; ++i) for(var j = 0; j != data[i].length; ++j) if(typeof data[i][j] === 'undefined') data[i][j] = "";

    // Loop over each row in the file
    for(var i=0; i < data.length; i++) {
        // add the file name to every row but the header
        if (i === 0) {
            data[i].unshift("BOM");                 // The first cell in header should say "BOM"
        } else {
            data[i].unshift(fileName.slice(0,-5));  // add the file name at the beginning without .xlsx
            data[i].push('\r\n');                   // add a line break at the end
        }

        var row = "";

        // output the properly formatted string (makes it work with comma data)
        for (var index in data[i]) {
            row += '"' + data[i][index] + '",';
        }

        // strip off the last trailing comma
        row.slice(0, row.length - 1);

        // Add proper line breaks
        row = row + '\r\n';

        // this loop skips over header rows after the first file
        if(loop > 1 && i !== 0 || loop === 1) {
            fs.appendFile(exportDirectory + exportFileName, row, function(err) {
                if(err) {
                    throw err;
                }
            })
        }
    }
}
