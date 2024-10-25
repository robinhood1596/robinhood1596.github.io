"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var videoStream = null;
function getCameraDevices() {
    return __awaiter(this, void 0, void 0, function () {
        var devices;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, navigator.mediaDevices.enumerateDevices()];
                case 1:
                    devices = _a.sent();
                    return [2 /*return*/, devices.filter(function (device) { return device.kind === 'videoinput'; })];
            }
        });
    });
}
function populateCameraDropdown() {
    return __awaiter(this, void 0, void 0, function () {
        var cameraDropdown, videoDevices;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    cameraDropdown = document.getElementById("cameraDropdown");
                    if (!cameraDropdown) {
                        console.error("Camera dropdown not found");
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, getCameraDevices()];
                case 1:
                    videoDevices = _a.sent();
                    videoDevices.forEach(function (device) {
                        var option = document.createElement("option");
                        option.value = device.deviceId;
                        option.text = device.label || "Camera ".concat(cameraDropdown.options.length + 1);
                        cameraDropdown.appendChild(option);
                    });
                    return [2 /*return*/];
            }
        });
    });
}
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
    return __awaiter(this, void 0, void 0, function () {
        var videoElement, canvasElement, captureButton, cameraDropdown, selectedDeviceId, stream, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    videoElement = document.querySelector("video");
                    canvasElement = document.querySelector("canvas");
                    if (!videoElement) {
                        videoElement = document.createElement("video");
                        document.body.appendChild(videoElement);
                    }
                    if (!canvasElement) {
                        canvasElement = document.createElement("canvas");
                        document.body.appendChild(canvasElement);
                    }
                    captureButton = document.getElementById("capturePhotoButton");
                    if (!captureButton) {
                        console.error("Capture button not found");
                        return [2 /*return*/];
                    }
                    if (!!videoStream) return [3 /*break*/, 4];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    cameraDropdown = document.getElementById("cameraDropdown");
                    if (!cameraDropdown) {
                        console.error("Camera dropdown not found");
                        return [2 /*return*/];
                    }
                    selectedDeviceId = cameraDropdown.value;
                    return [4 /*yield*/, navigator.mediaDevices.getUserMedia({ video: { deviceId: { exact: selectedDeviceId } } })];
                case 2:
                    stream = _a.sent();
                    videoStream = stream;
                    videoElement.srcObject = stream;
                    videoElement.play();
                    return [3 /*break*/, 4];
                case 3:
                    error_1 = _a.sent();
                    console.error("Error accessing the camera", error_1);
                    return [3 /*break*/, 4];
                case 4:
                    captureButton.onclick = function () {
                        if (videoStream) {
                            var desiredWidth = 1920; // Example width for high resolution
                            var desiredHeight = 1080; // Example height for high resolution
                            if (canvasElement) {
                                canvasElement.width = desiredWidth;
                                canvasElement.height = desiredHeight;
                                var context = canvasElement.getContext("2d");
                                if (context && videoElement) {
                                    context.drawImage(videoElement, 0, 0, desiredWidth, desiredHeight);
                                    var imageDataUrl = canvasElement.toDataURL("image/png").split(',')[1];
                                    console.log("Captured image data URL:", imageDataUrl);
                                    insertImageToSlide(imageDataUrl);
                                    context.clearRect(0, 0, canvasElement.width, canvasElement.height);
                                }
                                else {
                                    console.error("Failed to get canvas context or video element.");
                                }
                            }
                        }
                        else {
                            console.error("No video stream available");
                        }
                    };
                    return [2 /*return*/];
            }
        });
    });
}
function stopCamera() {
    if (videoStream) {
        videoStream.getTracks().forEach(function (track) { return track.stop(); });
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
    // Create and populate the camera dropdown
    var cameraDropdown = document.createElement("select");
    cameraDropdown.id = "cameraDropdown";
    document.body.appendChild(cameraDropdown);
    populateCameraDropdown();
});
