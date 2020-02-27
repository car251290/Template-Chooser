# Template-Chooser
Template that choose Word document.

#The second experience with Add-in

This Add-In and I made using Visual Studio and in the same language Javascript 
This case selected the document and display it in the Word document, To make a chooser it will be subtitle and they click the button and each of them will display templates and the user will select them, depends what type of document they are looking to display in Microsoft Word.

I code this function to display different subjects of templates in a list, also I code the view of the Add-in using the tools of CSS and HTML.

# Part of the CSS Code I use for making the display of the templates with a button and making the background 

# body of the page
I choose a color that was similiar to the document of Word so it will not be contraste and for the person using it will be easy to use.

# html 


# Javascript the next fuction is an example of the display template:

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



