# Template-Chooser
Template that choose Word document.

#The second experience with Add-in

This Add-In and I made using Visual Studio and in the same language Javascript 
This case selected the document and display it in the Word document, To make a chooser it will be subtitle and they click the button and each of them will display templates and the user will select them, depends what type of document they are looking to display in Microsoft Word.

I code this function to display different subjects of templates in a list, also I code the view of the Add-in using the tools of CSS and HTML.

#Part of the CSS Code I use for making the display of the templates with a button and making the background 
/*  ALL*/
body {
    margin: 0; /* Margin of the bar */
    font-family: Segoe UI, Segoe UI, serif; /* The style of the letter */
    text-align: left; /* Text align */
}
/* IconButton and text */
.btn {
    background-color: #f3f2f1; /* Blue background */
    border: none; /* Remove borders */
    color: #484644; /* black text */
    padding: 15px 23px; /* Some padding */
    cursor: pointer; /* Mouse pointer on hover */
    text-align: right; /* The text alination */
    left: 60px; /* Left of All the bar */
    width: 100%; /* Width of All the bar*/
    display: block; /* Block All */
    margin-top: 1px; /* Margin All the top */
}
    /* IconButton All and text */
    .btn :after {
        font-family: Segoe UI, Segoe UI, serif; /* Comment with the letter */
        content: 'ALL'; /* Text All of the bar */
        color: #000000;
        align-content: center; /* For center the text */
        visibility: visible; /* Visibility with of the text */
        position: absolute; /* Absolute in the bar */
        cursor: pointer; /* Cursor that pointer to the text */
        background-color: #f3f2f1; /* Color of the text */
        top: 2px 2px 2px 1px; /* Top of the text*/
        left: 20px; /* The text to the left */
        float: left; /* Float to left */
    }
    /* Darker background on mouse-over */
    .btn:hover {
        background-color: #f3f2f1; /* The color of the back ground when the button is display */
    }
/* The Item of All display */
#myDIV {
    width: 100%; /* The width of the Templates documents*/
    font-family: Segoe UI, Segoe UI, serif; /* Letter */
    padding: 12px 0px; /* The padding for the document and the bar */
    text-align: left; /* The alination of the document Templates*/
    margin-top: 1.5px; /* The margin it has */
    background-color: white; /* The background color */
    display: block; /* Display block */
}



