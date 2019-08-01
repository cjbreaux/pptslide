
(function () {
    "use strict";

    var messageBanner;

    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            $('#insert-title').click(insertTitle);
            $('#insert-content').click(insertContent);
            $('#title').keyup(insertPixabay);
            $('#jBold').click(function () {
                document.execCommand('bold');
                insertPixabay();
              
            });
           

            // Attach event handlers to images
            $(document).on('click', 'img', function () {
                var imgsrc = $(this).attr('src');
                toDataURL(imgsrc, function (dataUrl) {
                    insertImageFromBase64String(dataUrl)
                })
            });
        });
    };

    function insertTitle() {
        Office.context.document.setSelectedDataAsync($('#title').val(),
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function insertContent() {
        Office.context.document.setSelectedDataAsync($('#content').text(),
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    
    function insertPixabay() {
        var titleString = $('#title').val();
        var bold_words = $("strong").map(function () {
            return $(this).text();
        }).get();
       
        $.getJSON('https://pixabay.com/api/?key=13128735-e8d30252cb1da0dadc588a8d8&q=' + titleString + ' ' + bold_words + '&image_type=photo&per_page=12&pretty=true')
            .done(function (data) {
                $('#result').html(""); // Clear div
                $('img').off("click"); // Remove previous event handlers
                $.each(data.hits, function (key, pic) {
                    $('#result').append('<img src=' + pic.webformatURL + '</img>');                
                });
            });
    }

    function toDataURL(src, callback, outputFormat) {
        var img = new Image();
        img.crossOrigin = 'Anonymous';
        img.onload = function () {
            var canvas = document.createElement('CANVAS');
            var ctx = canvas.getContext('2d');
            var dataURL;
            canvas.height = this.naturalHeight;
            canvas.width = this.naturalWidth;
            ctx.drawImage(this, 0, 0);
            dataURL = canvas.toDataURL(outputFormat);
            dataURL = dataURL.substring(dataURL.indexOf(",") + 1);
            callback(dataURL);
        };
        img.src = src;
        if (img.complete || img.complete === undefined) {
            img.src = "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
            img.src = src;
        }
    }
   
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();