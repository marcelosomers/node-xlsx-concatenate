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

        // parse file
        to_json(file);
    }
});


function to_json(workbook) {
    var sheet_name_list = workbook.SheetNames;
    var sheet = workbook.Sheets[sheet_name_list[0]];
    var data = xlsx.utils.sheet_to_json(sheet, {header:1});

    fs.appendFile("output1.txt", data, function(err) {
        if(err) {
            throw err;
        }
    })
}
