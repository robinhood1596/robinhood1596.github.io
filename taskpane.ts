"use strict";

let videoStream: MediaStream | null = null;

function insertImageToSlide(imageDataUrl: string) {
    console.log("Inserting image.");
    Office.context.document.setSelectedDataAsync(
        imageDataUrl,
        { coercionType: Office.CoercionType.Image },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error inserting image: " + asyncResult.error.message);
            } else {
                console.log("Image inserted successfully.");
            }
        }
    );
}

function capturePhoto() {
    let videoElement = document.querySelector("video");
    let canvasElement = document.querySelector("canvas");

    if (!videoElement) {
        videoElement = document.createElement("video");
        document.body.appendChild(videoElement);
    }

    if (!canvasElement) {
        canvasElement = document.createElement("canvas");
        document.body.appendChild(canvasElement);
    }

    const captureButton = document.getElementById("capturePhotoButton");
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
            const desiredWidth = 1920; // Example width for high resolution
            const desiredHeight = 1080; // Example height for high resolution

            canvasElement.width = desiredWidth;
            canvasElement.height = desiredHeight;

            const context = canvasElement.getContext("2d");
            if (context) {
                context.drawImage(videoElement, 0, 0, desiredWidth, desiredHeight);
                const imageDataUrl = canvasElement.toDataURL("image/png").split(',')[1];
                console.log("Captured image data URL:", imageDataUrl);
                insertImageToSlide(imageDataUrl);

                context.clearRect(0, 0, canvasElement.width, canvasElement.height);
            } else {
                console.error("Failed to get canvas context.");
            }
        } else {
            console.error("No video stream available");
        }
    };
}

function stopCamera() {
    if (videoStream) {
        videoStream.getTracks().forEach(function (track) { track.stop(); });
        videoStream = null;
        const videoElement = document.querySelector("video");
        if (videoElement) {
            videoElement.remove();
        }
        const canvasElement = document.querySelector("canvas");
        if (canvasElement) {
            canvasElement.remove();
        }
    }
}

Office.onReady(function () {
    const captureButton = document.getElementById("capturePhotoButton");
    if (captureButton) {
        captureButton.addEventListener("click", capturePhoto);
    }

    const stopButton = document.createElement("button");
    stopButton.textContent = "Stop Camera";
    document.body.appendChild(stopButton);
    stopButton.addEventListener("click", stopCamera);
});