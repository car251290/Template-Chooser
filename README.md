# Template-Chooser
Templates for choose is to chooser word templates is storeted in SharePoint, the document display to the user choose.
In many Office 365 projects, SharePoint migrations or document structure planning, the question often araises, where to store templates so that every user has access to them and, more importantly, always uses the latest version.
Therefore, here is a tool powered approach that solves this problem and uses cloud technology to solve all these problems.

# The second experience with Add-in
This Add-In and I made using Visual Studio and in the same language Javascript 
This case selected the document and display it in the Word document, To make a chooser it will be subtitle and they click the button and each of them will display different templates there is a All button to display all of them in a single selection and easier the user will select them, depends what type of document they are looking to display in Microsoft Word.

I code this function to display different subjects of templates in a list, Also I code the view of the Add-in using the tools of CSS and HTML for the looking of the tempales.

# Part of the CSS Code I use for making the display of the templates with a button and making the background 
## Working in the Back end to connected to the Sharepoint is coming the code and the Template update it.

# body of the page
I choose a color that was similiar to the document of Word so it will not be contraste and for the person using it will be easy to use.

# Sideload Office Add-ins for testing from a network share
I will get this steps to deploy my Webpackes to Azure

Share a folder
In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.
Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose Properties.
Within the Properties dialog window, open the Sharing tab and then choose the Share button.

# html 
the body of the HTML has <script> where I indicate the src of the file of the js that will containe the html.
    <!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title></title>
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
    <script src="FunctionFile.js" type="text/javascript"></script>
</head>
<body>
  
</body>
</html>

# Javascript the next fuction is an example of the display template:

    function displaytemplates() {
        let templates = ['Templatechooser.docx', 'Template2.docx'];
        templates = new docxTemplater();
       templates.loadZip('zip');
        //forlook for the image
        for (let i = 0; i < templates.length; i++) {
            let File = templates[i];
            //add-in container for display the imagine with the url and the class html addin 
            $(".templates").append(
                '<div class= "tn">' +
                '<a" http://localhost/46TemplateChooserWeb/Templates' + File + '" alt = "templates" /> ' +
                '</div>'
            );
        }
    }



