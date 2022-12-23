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

app.get("/generatedDocs/Covering.docx", function(request, response){
    //  console.log(request);
    response.sendFile(__dirname+'/generatedDocs/Covering.docx');
    })

app.get("/generatedDocs/Undertaking.docx", function(request, response){
        //  console.log(request);
    response.sendFile(__dirname+'/generatedDocs/Undertaking.docx');
    })    

var fs = require('fs');
var path = require('path');

function generatedoc(docName, req){
    //Load the docx file as a binary
    var content = fs.readFileSync(path.resolve(__dirname+'/'+docName+'.docx'), 'binary');
    
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
            manufacturer_name: req.body.manufacturer_name,
            manufacturer_address: req.body.manufacturer_address,
            product_name: req.body.product_name,
            model_no:req.body.model_no,
            brand_name:req.body.brand_name,
            is_standard:req.body.is_standard,
            lab_name:req.body.lab_name,
            report_no:req.body.report_no,
            report_date:req.body.report_date,
            ir_name:req.body.ir_name,
            ir_address:req.body.ir_address
        });
    
        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
            errorHandler(error);
        }
        
        var buf = doc.getZip().generate({type: 'nodebuffer'});
        
        // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
        fs.writeFileSync(path.resolve(__dirname+'/generatedDocs/'+docName+'.docx'), buf);
    }

    app.post("/", function(req, res){
         generatedoc('Undertaking', req);
         generatedoc('Covering', req);


         res.sendFile(__dirname +"/downloadDocs.html");
          
    

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







app.listen(3000, function(){
    console.log("Server started on port 3000");
});


// const express = require('express')
// const app = express()
// app.all('/', (req, res) => {
//     console.log("Just got a request!")
//     res.send('Yo!')
// })
// app.listen(process.env.PORT || 3000)
