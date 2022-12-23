function generatedoc(docName){
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
    
    var buf = doc.getZip()
                 .generate({type: 'nodebuffer'});
    
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname+'/'+docName-filed+'.docx'), buf);
}

