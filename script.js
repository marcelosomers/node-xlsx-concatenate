var fs = require("fs")
    xlsx = require("xlsx");

var inputDirectory = "files/";

fs.readdir(inputDirectory, function(err, files) {
    if (err) {
        throw err;
    }

    // Loop over all files in the input directory
    for(var i=0, len=files.length; i < len; i++) {

        // Read the file
        var file = xlsx.readFile(inputDirectory + files[i]);

        // Parse file
        to_csv(file, files[i]);
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
}