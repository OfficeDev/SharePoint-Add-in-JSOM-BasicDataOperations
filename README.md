# Basic CRUD operations in SharePoint Add-ins using the JavaScript object model (JSOM) APIs #

## Summary
Use the SharePoint JavaScript object model (JSOM) to perform create, read, update, and delete operations on website properties, lists, and list items from a SharePoint Add-in.

### Applies to ###
-  SharePoint Online and on-premise SharePoint 2013 and later 

----------
## Prerequisites ##
This sample requires the following:


- A SharePoint 2013 development environment that is configured for app isolation. (A SharePoint Online Developer Site is automatically configured. For an on premise development environment, see [Set up an on-premises development environment for SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179923.aspx).) 


- Visual Studio and the Office Developer Tools for Visual Studio installed on your developer computer 


## Description of the code ##
The code that uses the JSOM APIs is located in the App.js and AppUIEvents.js file of the project. (The sample was created when SharePoint Add-ins were called "apps for SharePoint".) The top of the BasicTasksJSOM.html page of the add-in appears after you install and launch the add-in and looks similar to the following.

![The add-in start page with with links for viewing code, buttons for executing the code, and instructions for each pair.](/description/fig1.png) 



The sample demonstrates the following:


- How to read and write data to and from the add-in web of a SharePoint Add-in using the SharePoint JavaScript object model libray (JSOM).


- How to load the data returned from SharePoint into the client context object and then display the data. 


## To use the sample #

12. Open **Visual Studio** as an administrator.
13. Open the .sln file.
13. In **Solution Explorer**, highlight the SharePoint add-in project and replace the **Site URL** property with the URL of your SharePoint developer site.
14. Press F5. The add-in is installed and opens to its start page.
16. Each section of the start page describes the code for a programming task. There is a link that enables you to view the code and a button that enables you to execute the code. In some cases data appears right on the start page. In other cases, instructions appear that tell how to see the new or changed data.


## Questions and comments

We'd love to get your feedback on this sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/SharePoint-Add-in-JSOM-BasicDataOperations/issues) section of this repository.
  
<a name="resources"/>
## Additional resources

[SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179930.aspx)

[Complete basic operations using JavaScript library code](https://msdn.microsoft.com/library/office/jj163201.aspx)

### Copyright ###

Copyright (c) Microsoft. All rights reserved.






This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
