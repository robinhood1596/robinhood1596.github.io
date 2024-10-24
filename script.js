const video = document.getElementById('video');
const canvas = document.getElementById('canvas');
const context = canvas.getContext('2d');

navigator.mediaDevices.getUserMedia({ video: true })
  .then(stream => {
    video.srcObject = stream;
  })
  .catch(error => {
    console.error('Error accessing camera: ', error);
  });

document.getElementById('capture').addEventListener('click', () => {
  context.drawImage(video, 0, 0, canvas.width, canvas.height);
  const dataURL = canvas.toDataURL('image/png');

  Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
      insertImageInPowerPoint(dataURL);
    }
  });
});

function insertImageInPowerPoint(imageBase64) {
  PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getActiveSlide();
    slide.shapes.addImage(imageBase64, { left: 100, top: 100, width: 400, height: 300 });
    await context.sync();
  }).catch(error => {
    console.error("Error inserting image in PowerPoint: ", error);
  });
}
