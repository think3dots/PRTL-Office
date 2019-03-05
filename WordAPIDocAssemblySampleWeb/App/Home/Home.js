/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/// <reference path="../App.js" />

let doctorsArray = [];
let documentsArray = [];
let templateHeaders = [];
//localStorage.removeItem('DocsToSend');

(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            getHeader();
            getDoctorsList();
            // Hookup click events
            document.getElementById("send").onclick = function () {
                this.disabled = true;
                submitDocument();
            }
            // Configure Fabric NavBar
            $('.ms-NavBar').NavBar();
            $('#teamBuilder').click(function showTb() {
                $('#tbPanel').show();
                $('#sowPanel').hide();
            });
            $('#docGen').click(function showDg() {
                $('#tbPanel').hide();
                $('#sowPanel').show();
            });
            $('#docGen').click();
        });
    };
    

    /*****************************************************************************/
    /*
    /* Handle the results of each call to the service.
    /*
    /*****************************************************************************/

    function handleSuccess() {
        app.showNotification("Success", "Success");
        documentsArray = JSON.parse(localStorage.getItem('DocsToSend'));
        debugger
        if (documentsArray.lenght > 0) {
            var document = documentsArray[0];
            var formData = new FormData();

            formData.append('doc', document.fileContent);
            formData.append('referring', document.referring);
            formData.append('referral', document.referral);
            formData.append('email', document.email);
            formData.append('sms', document.sms);

            $.ajax({
                url: 'https://d538b722.ngrok.io/api/v1/documents?auth[key]=1234567890',
                type: 'POST',
                contentType: false,
                processData: false,
                data: formData
            }).done(function (data) {
                documentsArray.shift();
                localStorage.removeItem('DocsToSend');
                localStorage.setItem('DocsToSend', JSON.stringify(documentsArray));
                handleSuccess();
                app.showNotification("There are " + documentsArray.length + " remaining to be sent");
            }).fail(function () {
                app.showNotification('Error', 'Documents remaining to be sent: ' + documentsArray.length);
            }).always(function () {
                $('.disable-while-sending').prop('disabled', true);
            });
        }
        app.showNotification("All locally cached documents have been sent to the server");
        }

    function handleError() {
        app.showNotification("Error occurred in contacting the server");
    }

    function doctor(id, full_name) {
        this.id = id;
        this.full_name = full_name;
    }

    function serverDoc(fileContent, referring, referral, email, sms) {
        this.email = email;
        this.sms = sms;
        this.referring = referring;
        this.referral = referral;
        this.fileContent = fileContent;
    }

    function getDoctorsList() {
        doctorsArray = localStorage["doctors"];
        const localDoctorData = JSON.parse(localStorage.getItem('doctors'));
        var ds = localDoctorData.map(function (d) {
            return '<option value="' + d.id + '">' + d.full_name + '</option>';
        })
        $("select#referring-dropdown").html(ds.join(''));
        $("select#referral-dropdown").html(ds.join(''));

        $.ajax({
            url: 'https://d538b722.ngrok.io/api/v1/doctors?auth[key]=1234567890',
            type: 'GET',
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            if (JSON.stringify(data) !== doctorsArray) {
                localStorage.removeItem("doctors");
                var newDoctorArray = [];
                data.map(function (d) {
                    var newDoctor = new doctor(d.id, d.full_name)
                    newDoctorArray.push(newDoctor);
                })
                localStorage.setItem("doctors", JSON.stringify(newDoctorArray));
                var ds = newDoctorArray.map(function (d) {
                    return '<option value="' + d.id + '">' + d.full_name + '</option>';
                })
                $("select#referring-dropdown").html(ds.join(''));
                $("select#referral-dropdown").html(ds.join(''));
            } 
            }).fail(function () {
                handleError();
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }


    function submitDocument() {
        $('.disable-while-sending').prop('disabled', true);
        insertHeader();
        Office.context.document.getFileAsync(Office.FileType.Pdf,
            function (result) {
                if (result.status == "succeeded") { 
                    var myFile = result.value;
                    var sliceCount = myFile.sliceCount;
                    var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
                    app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
                    getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                    myFile.closeAsync();
                }
                else {
                    app.showNotification("Error:", result.error.message);
                }
            }
        );
    }

    function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status == "succeeded") {
                if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                    return;
                }
                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                if (++slicesReceived == sliceCount) {
                    // All slices have been received.
                    file.closeAsync();
                    onGotAllSlices(docdataSlices);
                }
                else {
                    getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                }
            }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }

    function onGotAllSlices(docdataSlices) {
        var docdata = [];
        var refr = document.getElementById('referring-dropdown').value;
        var refl = document.getElementById('referral-dropdown').value;
        var emailChecked = document.getElementById('send-email').checked;
        var SMSTicked = document.getElementById('send-sms').checked;
        
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);
        }

        var fileContent = new String();
        for (var j = 0; j < docdata.length; j++) {
            fileContent += String.fromCharCode(docdata[j]);
        }
        //encrypt filecontent for file transfer base64
        fileContent = btoa(fileContent);
        //Store the docData in case server is unavailable
        var newDocument = new serverDoc(fileContent, refr, refl, emailChecked, SMSTicked);
        // build the form that will be submitted to the server
        var formData = new FormData();
        documentsArray = JSON.parse(localStorage.getItem('DocsToSend')) || [];
        debugger
        documentsArray.push(newDocument);

        formData.append("doc", fileContent);
        formData.append("referring", refl);
        formData.append("referral", refr);
        formData.append("email", emailChecked);
        formData.append("sms", SMSTicked);

        $.ajax({
            url: 'https://d538b722.ngrok.io/api/v1/documents?auth[key]=1234567890',
            type: 'POST',
            contentType: false,
            processData: false,
            data: formData
        }).done(function (data) {
            documentsArray.shift();
            handleSuccess();
            }).fail(function () {
                localStorage.removeItem("DocsToSend");
                localStorage.setItem("DocsToSend", JSON.stringify(documentsArray));
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            $('.disable-while-sending').prop('disabled', true);
        });
    }

    function insertHeader() {
        Word.run(function (context) {
            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections;
            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');
            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                // Create a proxy object the primary header of the first section. 
                // Note that the header is a body object.
                var myHeader = mySections.items[0].body;
                // Queue a command to insert text at the end of the header.
                myHeader.insertInlinePictureFromBase64(headerPicture(), Word.InsertLocation.end);
                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log("Added a header to the first section.");
                });
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function getHeader() {
        var headerToDisplay = "";
        $.ajax({
            url: 'https://d538b722.ngrok.io/api/v1/templates?auth[key]=1234567890',
            type: 'GET',
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            headerToDisplay = data[0].header;
            localStorage.setItem("headerPicture", headerToDisplay);
            return headerToDisplay;
        }).fail(function (status) {
            return headerPicture();
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }

    function headerPicture() {
        templateHeaders = localStorage.getItem('headerPicture');
        return (templateHeaders);
    }

})();