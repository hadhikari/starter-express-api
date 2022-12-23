
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
const bodyParser=require("body-parser");
const express=require('express');
const app=express();
app.use(bodyParser.urlencoded({extended:true}));

app.get("/", function(request, response){
//  console.log(request);
response.sendFile(__dirname +"/index.html");
})

var fs = require('fs');
var path = require('path');

app.post("/", function(req, res){
    
    const nname=req.body.TAG1;
    const brand=req.body.TAG2;
    const is=req.body.TAG3;

    //Load the docx file as a binary
var content = fs
.readFileSync(path.resolve(__dirname+'/undertaking.docx'), 'binary');

var zip = new PizZip(content);
var doc;
try {
doc = new Docxtemplater(zip);
} catch(error) {
// Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
errorHandler(error);
}

//set the templateVariables
    doc.setData({
        first_name: nname,
        last_name: brand,
        phone: is
    });

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render()
    }
    catch (error) {
        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
        errorHandler(error);
    }
    
    var buf = doc.getZip()
                 .generate({type: 'nodebuffer'});
    
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname+'/undertaking-filed.docx'), buf);
    res.send("<h1>Hello,world!</h1>");

})

// The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
function replaceErrors(key, value) {
    if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
            error[key] = value[key];
            return error;
        }, {});
    }
    return value;
}

function errorHandler(error) {
    console.log(JSON.stringify({error: error}, replaceErrors));

    if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (error) {
            return error.properties.explanation;
        }).join("\n");
        console.log('errorMessages', errorMessages);
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
    }
    throw error;
}

//Load the docx file as a binary
var content = fs
    .readFileSync(path.resolve(__dirname+'/undertaking.docx'), 'binary');

var zip = new PizZip(content);
var doc;
try {
    doc = new Docxtemplater(zip);
} catch(error) {
    // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
    errorHandler(error);
}






app.listen(3000, function(){
    console.log("Server started on port 3000");
});
