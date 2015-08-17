// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.


// Website tasks
function retrieveWebsite(resultpanel) {
    var clientContext;

    clientContext = SP.ClientContext.get_current();
    this.oWebsite = clientContext.get_web();
    clientContext.load(this.oWebsite);

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Web site title: " + this.oWebsite.get_title();
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function retrieveWebsiteProps(resultpanel) {
    var clientContext;
    
    clientContext = new SP.ClientContext.get_current();
    this.oWebsite = clientContext.get_web();

    clientContext.load(this.oWebsite, "Description", "Created");

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Description: " + this.oWebsite.get_description() +
            "<br/>Date created: " + this.oWebsite.get_created();
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function writeWebsiteProps(resultpanel) {
    var clientContext;

    clientContext = new SP.ClientContext.get_current();
    this.oWebsite = clientContext.get_web();

    this.oWebsite.set_description("This is an updated description.");
    this.oWebsite.update();

    clientContext.load(this.oWebsite, "Description");

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );


    function successHandler() {
        resultpanel.innerHTML = "Web site description: " + this.oWebsite.get_description();
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

// Lists tasks
function readAllProps(resultpanel) {
    var clientContext;
    var oWebsite;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();

    this.collList = oWebsite.get_lists();
    clientContext.load(this.collList);

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listInfo;
        var listEnumerator;

        listEnumerator = this.collList.getEnumerator();
        
        listInfo = "";
        while (listEnumerator.moveNext()) {
            var oList = listEnumerator.get_current();
            listInfo += "Title: " + oList.get_title() + " Created: " +
                oList.get_created().toString() + "<br/>";
        }

        resultpanel.innerHTML = listInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function readSpecificProps(resultpanel) {
    var clientContext;
    var oWebsite;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();

    this.collList = oWebsite.get_lists();

    clientContext.load(this.collList, "Include(Title, Id)");

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listInfo;
        var listEnumerator;
        
        listEnumerator = this.collList.getEnumerator();

        listInfo = "";
        while (listEnumerator.moveNext()) {
            var oList = listEnumerator.get_current();
            listInfo += "Title: " + oList.get_title() +
                " ID: " + oList.get_id().toString() + "<br/>";
        }

        resultpanel.innerHTML = listInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function readColl(resultpanel) {
    var clientContext;
    var oWebsite;
    var collList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    collList = oWebsite.get_lists();

    this.listInfoCollection = clientContext.loadQuery(collList, "Include(Title, Id)");

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listInfo;

        listInfo = "";
        for (var i = 0; i < this.listInfoCollection.length; i++) {
            var oList = this.listInfoCollection[i];
            listInfo += "Title: " + oList.get_title() +
                " ID: " + oList.get_id().toString() + "<br/>";
        }
        
        resultpanel.innerHTML = listInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function readFilter(resultpanel) {
    var clientContext;
    var oWebsite;
    var collList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    collList = oWebsite.get_lists();

    this.listInfoArray = clientContext.loadQuery(collList,
        "Include(Title,Fields.Include(Title,InternalName))");

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listInfo;

        for (var i = 0; i < this.listInfoArray.length; i++) {
            var oList = this.listInfoArray[i];
            var collField = oList.get_fields();
            var fieldEnumerator = collField.getEnumerator();

            listInfo = "";
            while (fieldEnumerator.moveNext()) {
                var oField = fieldEnumerator.get_current();
                var regEx = new RegExp("name", "ig");

                if (regEx.test(oField.get_internalName())) {
                    listInfo += "List: " + oList.get_title() +
                        "<br/>&nbsp;&nbsp;&nbsp;&nbsp;Field Title: " + oField.get_title() +
                        "<br/>&nbsp;&nbsp;&nbsp;&nbsp;Field Internal name: " + oField.get_internalName();
                }
            }
        }

        resultpanel.innerHTML = listInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

// Create, update and delete lists
function createList(resultpanel) {
    var clientContext;
    var oWebsite;
    var listCreationInfo;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();

    listCreationInfo = new SP.ListCreationInformation();
    listCreationInfo.set_title("My Announcements List");
    listCreationInfo.set_templateType(SP.ListTemplateType.announcements);

    this.oList = oWebsite.get_lists().add(listCreationInfo);
    clientContext.load(this.oList);

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/My Announcements List'>list</a>.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function updateList(resultpanel) {
    var clientContext;
    var oWebsite;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();

    this.oList = oWebsite.get_lists().getByTitle("My Announcements List");
    this.oList.set_description("New Announcements List");
    this.oList.update();

    clientContext.load(this.oList);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Check the description in the <a href='../Lists/My Announcements List'>list</a>.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function addField(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;
    var fieldNumber;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("My Announcements List");

    this.oField = oList.get_fields().addFieldAsXml(
        "<Field DisplayName='MyField' Type='Number' />",
        true,
        SP.AddFieldOptions.defaultValue
    );

    fieldNumber = clientContext.castTo(this.oField, SP.FieldNumber);
    fieldNumber.set_maximumValue(100);
    fieldNumber.set_minimumValue(35);
    fieldNumber.update();

    clientContext.load(this.oField);

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "The <a href='../Lists/My Announcements List'>list</a> with a new field.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function deleteList(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;

    this.listTitle = "My Announcements List";

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle(this.listTitle);
    oList.deleteObject();

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = this.listTitle + " deleted.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

// Create, update and delete folders
function createFolder(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;
    var itemCreateInfo;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Shared Documents");

    itemCreateInfo = new SP.ListItemCreationInformation();
    itemCreateInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
    itemCreateInfo.set_leafName("My new folder!");
    this.oListItem = oList.addItem(itemCreateInfo);
    this.oListItem.update();

    clientContext.load(this.oListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Shared Documents'>document library</a> to see your new folder.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function updateFolder(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Shared Documents");

    this.oListItem = oList.getItemById(1);
    this.oListItem.set_item("FileLeafRef", "My updated folder");
    this.oListItem.update();

    clientContext.load(this.oListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Shared Documents'>document library</a> to see your updated folder.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function deleteFolder(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Shared Documents");

    this.oListItem = oList.getItemById(1);
    this.oListItem.deleteObject();

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Shared Documents'>document library</a> to make sure the folder is no longer there.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

// List item tasks
function readItems(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;
    var camlQuery;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Announcements");
    camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query><Where><Geq><FieldRef Name=\'ID\'/>' +
        '<Value Type=\'Number\'>1</Value></Geq></Where></Query>' +
        '<RowLimit>10</RowLimit></View>'
    );
    this.collListItem = oList.getItems(camlQuery);

    clientContext.load(this.collListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listItemInfo;
        var listItemEnumerator;

        listItemEnumerator = this.collListItem.getEnumerator();

        listItemInfo = "";
        while (listItemEnumerator.moveNext()) {
            var oListItem;
            oListItem = listItemEnumerator.get_current();
            listItemInfo += "ID: " + oListItem.get_id() + "<br/>" +
                "Title: " + oListItem.get_item("Title") + "<br/>" +
                "Body: " + oListItem.get_item("Body") + "<br/>";
        }

        resultpanel.innerHTML = listItemInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function readInclude(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;
    var camlQuery;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Announcements");
    camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><RowLimit>100</RowLimit></View>');

    this.collListItem = oList.getItems(camlQuery);

    clientContext.load(this.collListItem, "Include(Id, DisplayName, HasUniqueRoleAssignments)");
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        var listItemInfo;
        var listItemEnumerator;

        listItemEnumerator = this.collListItem.getEnumerator();

        listItemInfo = "";
        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            listItemInfo += "ID: " + oListItem.get_id() + "<br/>" +
            "Display name: " + oListItem.get_displayName() + "<br/>" +
            "Unique role assignments: " + oListItem.get_hasUniqueRoleAssignments() + "<br/>";
        }

        resultpanel.innerHTML = listItemInfo;
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

// Create, update and delete list items
function createListItem(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;
    var itemCreateInfo;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Announcements");

    itemCreateInfo = new SP.ListItemCreationInformation();
    this.oListItem = oList.addItem(itemCreateInfo);
    this.oListItem.set_item("Title", "My New Item!");
    this.oListItem.set_item("Body", "Hello World!");
    this.oListItem.update();
    
    clientContext.load(this.oListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Announcements'>list</a> to see your new item.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function updateListItem(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Announcements");

    this.oListItem = oList.getItemById(1);
    this.oListItem.set_item("Title", "My updated title");
    this.oListItem.update();

    clientContext.load(this.oListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Announcements'>list</a> to see your updated item.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

function deleteListItem(resultpanel) {
    var clientContext;
    var oWebsite;
    var oList;

    clientContext = new SP.ClientContext.get_current();
    oWebsite = clientContext.get_web();
    oList = oWebsite.get_lists().getByTitle("Announcements");

    this.oListItem = oList.getItemById(1);
    this.oListItem.deleteObject();

    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        resultpanel.innerHTML = "Go to the <a href='../Lists/Announcements'>list</a> to make sure the item is no longer there.";
    }

    function errorHandler() {
        resultpanel.innerHTML = "Request failed: " + arguments[1].get_message();
    }
}

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