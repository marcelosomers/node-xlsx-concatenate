var fs = require("fs")
    xlsx = require("xlsx");

var inputDirectory = "files/";
var csvContent = "";

fs.readdir(inputDirectory, function(err, files) {
    if (err) {
        throw err;
    }

    // Loop over all files in the input directory
    for(var i=0, len=files.length; i < len; i++) {

        var fileName = files[i];

        // TODO: Only open xlsx files
        if(files[i] !== ".DS_Store") {
            // Read the file
            var file = xlsx.readFile(inputDirectory + files[i]);

            // parse file
            to_json(file, fileName, i);
        }
    }
});


function to_json(workbook, fileName, loop) {
    var sheet_name_list = workbook.SheetNames;
    var sheet = workbook.Sheets[sheet_name_list[0]];
    var data = xlsx.utils.sheet_to_json(sheet, {header:1});

    // Process the blank cells to use an empty string
    for(var i = 0; i != data.length; ++i) for(var j = 0; j != data[i].length; ++j) if(typeof data[i][j] === 'undefined') data[i][j] = "";

    // Loop over each row in the file
    for(var i=0; i<data.length; i++) {
        // add the file name to every row but the header
        if (i === 0) {
            data[i].unshift("BOM")      // The first cell in header should say "BOM"
        } else {
            data[i].unshift(fileName);  // add the file name at the beginning
            data[i].push('\r\n');       // add a line break at the end
        }

        var row = "";

        // output the properly formatted string (makes it work with comma data)
        for (var index in data[i]) {
            row += '"' + data[i][index] + '",';
        }

        // strip off the last trailing comma
        row.slice(0, row.length - 1);

        if(loop !== 0 && row.indexOf("Part No.") > -1) {

        } else {
            csvContent += row + '\r\n';
        }
    }


    // Write to output1.csv
    fs.appendFile("output1.csv", csvContent, function(err) {
        if(err) {
            throw err;
        }
    })
}