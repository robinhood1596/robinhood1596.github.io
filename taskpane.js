!function(){"use strict";var e={81192:function(e,t,o){e.exports=o.p+"35e04bc3a9634afe0d10.js"}},t={};function o(n){var a=t[n];if(void 0!==a)return a.exports;var r=t[n]={exports:{}};return e[n](r,r.exports,o),r.exports}o.m=e,o.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),o.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;o.g.importScripts&&(e=o.g.location+"");var t=o.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var a=n.length-1;a>-1&&(!e||!/^http(s?):/.test(e));)e=n[a--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),o.p=e}(),o.b=document.baseURI||self.location.href,function(){let e=null;function t(e){const t=document.getElementById("statusMessage");t&&(t.textContent=e)}async function o(e){const t=e.getCapabilities();if(e.getSettings(),t.focusMode&&Array.isArray(t.focusMode)&&t.focusMode.includes("continuous"))try{await e.applyConstraints({advanced:[{focusMode:"continuous"}]})}catch(e){console.warn("Failed to set auto-focus:",e)}}async function n(){let n=document.querySelector("video");if(n||(n=document.createElement("video"),document.body.appendChild(n)),document.getElementById("capturePhotoButton")){if(!e)try{const a=document.getElementById("cameraDropdown"),r=document.getElementById("resolutionDropdown");if(!a||!r)return void console.error("Dropdowns not found");const c=a.value,[s,i]=r.value.split("x").map(Number);t("Accessing camera...");const u=await navigator.mediaDevices.getUserMedia({video:{deviceId:{exact:c},width:{ideal:s},height:{ideal:i},advanced:[{focusMode:{ideal:"continuous"}}]}});e=u,n.srcObject=u,await n.play();const d=e.getVideoTracks()[0];let l=!1;try{l=await async function(e){const t=e.getCapabilities();return t.focusMode&&Array.isArray(t.focusMode)&&t.focusMode.includes("continuous")}(d)}catch(e){console.warn("Error checking auto-focus support:",e)}if(l){t("Camera ready. Adjusting focus...");try{await o(d),await new Promise((e=>setTimeout(e,2e3))),t("Auto-focus set. Ready to capture!")}catch(e){console.warn("Error setting auto-focus:",e),t("Auto-focus not available. Ready to capture!")}}else t("Auto-focus not supported. Ready to capture!")}catch(e){console.error("Error accessing the camera",e),t("Error accessing the camera. Please try again.")}if(e){t("Capturing photo...");const o=e.getVideoTracks()[0],n=new ImageCapture(o);try{const e=await n.takePhoto(),o=await(a=e,new Promise(((e,t)=>{const o=new FileReader;o.onload=()=>e(o.result),o.onerror=t,o.readAsDataURL(a)})));console.log("Captured image data URL:",o),function(e){Office.context.document.setSelectedDataAsync(e,{coercionType:Office.CoercionType.Image},(function(e){e.status===Office.AsyncResultStatus.Failed?(console.error("Error inserting image: "+e.error.message),t("Failed to insert image. Please try again.")):(console.log("Image inserted successfully."),t("Image inserted successfully!"))}))}(o.split(",")[1]),t("Photo captured and inserted!")}catch(e){console.error("Error capturing photo",e),t("Error capturing photo. Please try again.")}}else console.error("No video stream available"),t("Camera not ready. Please start the camera first.");var a}else console.error("Capture button not found")}function a(){if(e){e.getTracks().forEach((e=>e.stop())),e=null;const o=document.querySelector("video");o&&o.remove(),t("Camera stopped.")}}async function r(){if(e){const n=e.getVideoTracks()[0];t("Readjusting focus...");try{await o(n),await new Promise((e=>setTimeout(e,2e3))),t("Focus readjusted!")}catch(e){console.warn("Error readjusting focus:",e),t("Failed to readjust focus. Please try again.")}}else console.error("No video stream available"),t("Camera not ready. Please start the camera first.")}Office.onReady((function(){const e=document.getElementById("capturePhotoButton");e&&e.addEventListener("click",n);const t=document.getElementById("stopCameraButton");t&&t.addEventListener("click",a);const o=document.createElement("button");o.textContent="Readjust Focus",o.id="readjustFocusButton",document.body.appendChild(o),o.addEventListener("click",r);const c=document.createElement("select");c.id="cameraDropdown",document.body.appendChild(c),async function(){const e=document.getElementById("cameraDropdown");(await navigator.mediaDevices.enumerateDevices()).filter((e=>"videoinput"===e.kind)).forEach(((t,o)=>{const n=document.createElement("option");n.value=t.deviceId,n.text=t.label||`Camera ${o+1}`,e.appendChild(n)}))}();const s=document.createElement("select");s.id="resolutionDropdown",document.body.appendChild(s),function(){const e=document.getElementById("resolutionDropdown");[{width:640,height:480,label:"VGA"},{width:1280,height:720,label:"HD"},{width:1920,height:1080,label:"Full HD"},{width:3712,height:2784,label:"4K"}].forEach((t=>{const o=document.createElement("option");o.value=`${t.width}x${t.height}`,o.text=`${t.label} (${t.width}x${t.height})`,e.appendChild(o)}))}();const i=document.createElement("div");i.id="statusMessage",document.body.appendChild(i)}))}(),new URL(o(81192),o.b)}();
//# sourceMappingURL=taskpane.js.map