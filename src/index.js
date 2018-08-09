'use strict';

(function () {

    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Set up event handler for the UI.
            $('#addProject').click(addProjectToPage);
        });
    };

    function ShouldIDisplayThis(p){
        if ($("#cbSystem").prop("checked"))
            return true;
        if (p.startsWith("System.") && p!="System.Description" || p.startsWith("WEF_"))
            return false;
        else
            return true;
    }
    function ShouldIDisplayThisActivityField(p){
        if ($("#cbSystem").prop("checked"))
            return true;
        if (p.startsWith("WEF_") || p=="System.IterationPath" || p=="System.WorkItemType")
            return false;
        else
            return true;
    }

    function CleanTitle(p){
        if ($("#cbCleanTitles").prop("checked")){
            var rP=p.replace("CSEngineering-V2.","");
            rP=rP.replace("CSEngineering-V2-Orgs.","Orgs.")
            rP=rP.replace("System.","")
            return rP;
        }
        else
            return p;
    }

    function addProjectToPage() {        
        OneNote.run(function (context) {
            // your collection url
            var collectionUrl = "https://csefy19.visualstudio.com/_apis/wit/workItems/";
            // ideally from config
            var token = $("#textToken").val();
            var request = require("request");
            var encodedPat = encodePat(token);
             
            var options = {
               method: 'GET',
               headers: { 'cache-control': 'no-cache', 'authorization': `Basic ${encodedPat}` },
               url: collectionUrl + $('#textProject').val(),
               qs: { 'api-version': '4.1' }
            };
             
            request(options, function (error, response, body) {
               if (error) {
                   var sError="ERROR<br>Token:" + $("#textToken").val() + "<br>Message:" + JSON.stringify(error.message);
                   var page = context.application.getActivePage();
                   // Queue a command to load the page with the title property.             
                   page.load('title'); 
                   // Add an outline with the specified HTML to the page.
                   var outline = page.addOutline(80, 120, sError);
                   context.sync();
               }
               

               var res=JSON.parse(response.body);

               var html = "<b>" + res.fields["System.WorkItemType"] + "</b> "+ $("#textProject").val() + 
                        "<br/>" + 
                        "<table border=1>" 
                switch (res.fields["System.WorkItemType"]) {
                    case "Activity":
                        for(var p in res.fields){
                            if (ShouldIDisplayThisActivityField(p))
                            html+="<tr><td>" + CleanTitle(p) + "</td><td>" + res.fields[p] + "</td></tr>";
                        }
                    break;
                    default:
                        for(var p in res.fields){
                            if (ShouldIDisplayThis(p))
                            html+="<tr><td>" + CleanTitle(p) + "</td><td>" + res.fields[p] + "</td></tr>";
                        }
                        break;
                }
               html+="</table>";

               // Get the current page.
               var page = context.application.getActivePage();
               // Queue a command to load the page with the title property.             
               page.load('title'); 
               // Add an outline with the specified HTML to the page.
               var outline = page.addOutline(80, 120, html);
               context.sync();
             });
            console.log("done");
             
            function encodePat(pat) {
               var b = new Buffer(':' + pat);
               var s = b.toString('base64');
             
               return s;
            }



            var html = '<table border=1><tr><td>Called VSTS</td></tr><tr><td>' + res + '</td><td><p>' + $('#textBox').val() + '</p></td></tr></table>';

            // Get the current page.
            var page = context.application.getActivePage();

            // Queue a command to load the page with the title property.             
            page.load('title'); 

            // Add an outline with the specified HTML to the page.
            var outline = page.addOutline(40, 90, html);

            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function() {
                    console.log('Added outline to page ' + page.title);
                })
                .catch(function(error) {
                    app.showNotification("Error: " + error); 
                    console.log("Error: " + error); 
                    if (error instanceof OfficeExtension.Error) { 
                        console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                    } 
                }); 
        });
    }

})();