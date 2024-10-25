"use strict";
var videoStream = null;
function insertImageToSlide(imageDataUrl) {
    console.log("Inserting image.");
    Office.context.document.setSelectedDataAsync(imageDataUrl, { coercionType: Office.CoercionType.Image }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Error inserting image: " + asyncResult.error.message);
        }
        else {
            console.log("Image inserted successfully.");
        }
    });
}
function capturePhoto() {
    var videoElement = document.querySelector("video");
    var canvasElement = document.querySelector("canvas");
    if (!videoElement) {
        videoElement = document.createElement("video");
        document.body.appendChild(videoElement);
    }
    if (!canvasElement) {
        canvasElement = document.createElement("canvas");
        document.body.appendChild(canvasElement);
    }
    var captureButton = document.getElementById("capturePhotoButton");
    if (!captureButton) {
        console.error("Capture button not found");
        return;
    }
    if (!videoStream) {
        navigator.mediaDevices.getUserMedia({ video: true })
            .then(function (stream) {
            videoStream = stream;
            videoElement.srcObject = stream;
            videoElement.play();
        })
            .catch(function (error) {
            console.error("Error accessing the camera", error);
        });
    }
    captureButton.onclick = function () {
        if (videoStream) {
            var desiredWidth = 1920; // Example width for high resolution
            var desiredHeight = 1080; // Example height for high resolution
            canvasElement.width = desiredWidth;
            canvasElement.height = desiredHeight;
            var context = canvasElement.getContext("2d");
            if (context) {
                context.drawImage(videoElement, 0, 0, desiredWidth, desiredHeight);
                var imageDataUrl = canvasElement.toDataURL("image/png").split(',')[1];
                console.log("Captured image data URL:", imageDataUrl);
                insertImageToSlide(imageDataUrl);
                context.clearRect(0, 0, canvasElement.width, canvasElement.height);
            }
            else {
                console.error("Failed to get canvas context.");
            }
        }
        else {
            console.error("No video stream available");
        }
    };
}
function stopCamera() {
    if (videoStream) {
        videoStream.getTracks().forEach(function (track) { track.stop(); });
        videoStream = null;
        var videoElement = document.querySelector("video");
        if (videoElement) {
            videoElement.remove();
        }
        var canvasElement = document.querySelector("canvas");
        if (canvasElement) {
            canvasElement.remove();
        }
    }
}
Office.onReady(function () {
    var captureButton = document.getElementById("capturePhotoButton");
    if (captureButton) {
        captureButton.addEventListener("click", capturePhoto);
    }
    var stopButton = document.createElement("button");
    stopButton.textContent = "Stop Camera";
    document.body.appendChild(stopButton);
    stopButton.addEventListener("click", stopCamera);
});
