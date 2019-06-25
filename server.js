var xlsx = require("xlsx");
var fs = require("fs");
var path = require("path");
var sourceDir = "Files";
var _ = require("lodash");




function filetojson(file){
    var wb = xlsx.readFile(file);
    var firsttablename = wb.SheetNames[0];
    var ws = wb.Sheets[firsttablename];
    var data = xlsx.utils.sheet_to_json(ws);
    return data;
}
var targetDir = path.join(__dirname,"Files");
var files = fs.readdirSync(targetDir);

var combinedData = [];

files.forEach(function(file){
    var fileext = path.parse(file).ext;
    if(fileext ==".xlsx"){
        var fullFilePath = path.join(__dirname,sourceDir,file);       
        var data = filetojson(fullFilePath);       
        combinedData = combinedData.concat(data);
    }
});
var newData = [];

newData =_.uniqWith(combinedData,_.isEqual);
var newWb = xlsx.utils.book_new();
var newWs = xlsx.utils.json_to_sheet(newData);
xlsx.utils.book_append_sheet(newWb,newWs,"Combined Data");

xlsx.writeFile(newWb,"newcombineddata.xlsx");
console.log("Done");

