(function () {
    "use strict";

    var messageBanner;

    Office.onReady(function () {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // TODO1: Assign event handler for insert-image button.
            $('#insert-image').click(insertImage);
            // TODO4: Assign event handler for insert-text button.
            $('#insert-text').click(insertText);

            // TODO6: Assign event handler for get-slide-metadata button.
            $('#get-slide-metadata').click(getSlideMetadata);
            // TODO8: Assign event handlers for the four navigation buttons.
            $('#go-to-first-slide').click(goToFirstSlide);
            $('#go-to-next-slide').click(goToNextSlide);
            $('#go-to-previous-slide').click(goToPreviousSlide);
            $('#go-to-last-slide').click(goToLastSlide);
            $('#insert-file').click(insertFile);
            $('#file').change(storeFileAsBase64)
            $('#insert-64file').change(insertAfterSelectedSlide)
        });
    });
    //wayne 10/15
   

    async function insertAllSlides() {
        await PowerPoint.run(async function (context) {
            context.presentation.insertSlidesFromBase64(chosenFileBase64);
            await context.sync();
        });
    }
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function (context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    function insertFile() {
        // Get file from from web service (as a Base64 encoded string).
        $.ajax({
            url: "NETORGFT6819296.onmicrosoft.com", success: function (result) {
                insertFileFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    function insertFileFromBase64String(File) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(File, {
            coercionType: Office.CoercionType.File
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    // TODO2: Define the insertImage function. 
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }

    // TODO3: Define the insertImageFromBase64String function.
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
    // TODO5: Define the insertText function.
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    // TODO7: Define the getSlideMetadata function.
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    // TODO9: Define the navigation functions.
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();