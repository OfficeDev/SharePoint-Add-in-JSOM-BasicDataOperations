// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.


$(":button.code-exec").click(function () {
    $("#" + this.parentElement.id + " .result-panel").toggle();

    if ($("#" + this.parentElement.id + " .result-panel").is(":visible")) {
        var funcName;
        var func;


        funcName = this.parentElement.id.replace("Container", "");
        func = window[funcName];
        func($("#" + this.parentElement.id + " .result-panel")[0]);
    }
});

$("a.code-link").click(function () {
    $("#" + this.parentElement.id + " .code-content").toggle();

    if ($("#" + this.parentElement.id + " .code-content").is(":visible")) {
        var funcName;
        var funcText;

        funcName = this.parentElement.id.replace("Container", "");
        funcText = window[funcName].toString();
        funcText = $("<div></div>").text(funcText).html();
        funcText = funcText.replace(/\r\n/g, "<br/>");
        funcText = funcText.replace(/ /g, "&nbsp;");

        $("#" + this.parentElement.id + " .code-content").html(funcText)
    }
});

/*

SharePoint-Add-in-JSOM-BasicDataOperations, https://github/officedev/SharePoint-Add-in-JSOM-BasicDataOperations
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/