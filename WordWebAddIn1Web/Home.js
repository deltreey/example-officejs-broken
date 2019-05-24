
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            Word.run(function (context) {
                context.document.body.clear();
                var cc = context.document.body
                    .insertParagraph(
                        "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
                        "Start"
                    )
                    .insertContentControl();
                context.sync().then(function() {
                    setInterval(readDocument, 5000);
                });
            });
        });
    };

    function readDocument() {
        Word.run(function (context) {
            var contentControls = context.document.body.contentControls;
            context.load(contentControls, 'title');
            return context.sync().then(function () {
                clearContent();
                if (contentControls.items.length === 0) {
                    writeContent("No Content Controls found.");
                    return;
                }
                var ooxmlContent = [];
                for (var i = 0; i < contentControls.items.length; ++i) {
                    writeContent(contentControls.items[i].title);
                    // reading the ooxml causes the UI to iterate over every content control
                    ooxmlContent.push(contentControls.items[i].getOoxml());
                }
                context.sync().then(function () {
                    for (var j = 0; j < ooxmlContent.length; ++j) {
                        var contentText = ooxmlContent[j].value;
                        console.log(contentText);
                    }
                });
            });
        })
        .catch (function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }

    function writeContent(text) {
        document.getElementById("message").innerHTML += "<br>" + text;
    }

    function clearContent() {
        document.getElementById("message").innerHTML = "";
    }
})();
