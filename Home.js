
(function () {
    "use strict";
    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            displaytemplates();
            generate();
            // Use this to check whether the API is supported in the Word client.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }
        });

        // Do something that is only available via the new APIs
        //Selection of image, insert image
        $('.tn').on('click', function (event) {
            var images = event.currentTarget.querySelector("img");
            var url = images.src;
            console.log("insert Templates");
            //to insert the image from the function!
            toDataURL(url, function (dataUrl) {
                insertTemplate(dataUrl);

            });
        });

    });


    //Function for display the imagine in the addin
    function displaytemplates() {
        var templates = ['Templatechooser.docx', 'Template2.docx'];
        templates = new docxTemplater();
       templates.loadZip(zip);
        //forlook for the image
        for (var i = 0; i < templates.length; i++) {
            var File = templates[i];
            //add-in container for display the imagine with the url and the class html addin 
            $(".templates").append(
                '<div class= "tn">' +
                '<img src=" http://localhost/46TemplateChooserWeb/Images/' + File + '" alt = "templates" > ' +
                '</div>'
            );
        }
    }
    
   

    // toDataUrl fuction to get the date of the image
    function toDataURL(url, callback) {
        // try to get the binary of the document

        templates.DocxTemplater.getBinaryContent(url, callback)
        console.log('template document')
        //method for the request of the data
        var xhr = new XMLHttpRequest();
        xhr.onload = function () {
            var reader = new FileReader();
            reader.onloadend = function () {
                callback(reader.result.split(',')[1]);
            }
            reader.readAsDataURL(xhr.response);
        };
        //to open the url and get the data of the image selected
        xhr.open('GET', url);
        xhr.responseType = 'blob';
        xhr.send();
        console.log('toDataURL');
    }

    function insertTemplate(base64) {
        Word.run(function (context) {
            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            //var range = context.document.getSelection();
            // Queue a command to replace the selected text.
           // range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            // Create a proxy object for the document body.
            // var body = context.document.body;
            var body = context.document.getSelection();

            // Queue a command to insert base64 encoded .docx at the beginning of the content body.
            // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.

             body.insertFileFromBase64(getBase64(), Word.InsertLocation.replace);

            //body.insertFileFromBase64(base64, Word.InsertLocation.replace);
            return context.sync().then(function () {
                console.log('Added template.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });

   
    }
    // Disaple the zoon of the add-in
    $(document).keydown(function (event) {
        if (event.ctrlKey == true && (event.which == '61' || event.which == '107' || event.which == '173' || event.which == '109' || event.which == '187' || event.which == '189')) {
            event.preventDefault();
        }
        // 107 Num Key  +
        // 109 Num Key  -
        // 173 Min Key  hyphen/underscor Hey
        // 61 Plus key  +/= key
    });
    $(window).bind('mousewheel DOMMouseScroll', function (event) {
        if (event.ctrlKey == true) {
            event.preventDefault();
        }
    });

    // try to get the document and display it
    var PizZip = require('Template2.docx');
    var Templatedoc = require('Template2.docx');

    var fs = require('fs');
    var path = require('path');

    //Load the docx file as a binary
    var content = fs
        .readFileSync(path.resolve(__dirname, 'Template2.docx'), 'binary');

    var zip = new PizZip(content);

    var doc = new Templatedoc();
    doc.loadZip(Template2.docx);

    //set the templateVariables
    doc.setData({
        first_name: 'Javier',
        last_name: 'carreno',
        phone: '0652455478',
        description: 'New Website'
    });

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render()
    }
    catch (error) {
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
        function replaceErrors(key, value) {
            if (value instanceof Error) {
                return Object.getOwnPropertyNames(value).reduce(function (error, key) {
                    error[key] = value[key];
                    return error;
                }, {});
            }
            return value;
        }
        console.log(JSON.stringify({ error: error }, replaceErrors));

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
    // try to get the scr of the document 
   // scr = "http://localhost/46TemplateChooserWeb/Images/"
    var buf = doc.getZip()
        .generate({ type: 'nodebuffer' });

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);

    function loadFile(url, callback) {
        PizZipUtils.getBinaryContent(url, callback);
    }
    function generate() {
        loadFile("http://localhost/46TemplateChooserWeb/Images/", function (error, content) {
            if (error) { throw error };
            var zip = new PizZip(content);
            var doc = new window.Templatedoc().loadZip(zip)
            doc.setData({
                first_name: 'Javier',
                last_name: 'carreno',
                phone: '0652455478',
                description: 'New Website'
            });
            try {
                // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                doc.render()
            }
            catch (error) {
                // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
                function replaceErrors(key, value) {
                    if (value instanceof Error) {
                        return Object.getOwnPropertyNames(value).reduce(function (error, key) {
                            error[key] = value[key];
                            return error;
                        }, {});
                    }
                    return value;
                }
                console.log(JSON.stringify({ error: error }, replaceErrors));

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
            var out = doc.getZip().generate({
                type: "blob",
                mimeType: "http://localhost/46TemplateChooserWeb/Images/"
            }) //Output the document using Data-URI
            saveAs(out, "output.docx")
        })
    }
   
})();
