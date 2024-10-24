let videoStream;

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Office ist bereit.");
        setupCamera();
    } else {
        console.error("Nicht in PowerPoint.");
    }
});

function setupCamera() {
    const video = document.getElementById('video');
    const canvas = document.getElementById('canvas');
    const context = canvas.getContext('2d');

    console.log("Versuche, die Kamera zu aktivieren...");

    navigator.mediaDevices.getUserMedia({ video: true })
        .then(stream => {
            console.log("Kamera erfolgreich aktiviert.");
            videoStream = stream;
            video.srcObject = stream;
        })
        .catch(error => {
            console.error('Fehler beim Zugriff auf die Kamera: ', error);
            alert("Fehler beim Zugriff auf die Kamera: " + error.message);
        });

    document.getElementById('capture').addEventListener('click', () => {
        console.log("Foto wird erfasst...");
        context.drawImage(video, 0, 0, canvas.width, canvas.height);
        const dataURL = canvas.toDataURL('image/png');
        console.log("BilddatenURL: ", dataURL);
        insertImageInPowerPoint(dataURL);
    });
}

function insertImageInPowerPoint(imageBase64) {
    console.log("Versuche, Bild in PowerPoint einzufügen...");
    PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getActiveSlide();
        slide.shapes.addImage(imageBase64, { left: 100, top: 100, width: 400, height: 300 });
        await context.sync();
        console.log("Bild erfolgreich in PowerPoint eingefügt.");
    }).catch(error => {
        console.error("Fehler beim Einfügen des Bildes in PowerPoint: ", error);
    });
}
