function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.logoDataUri = getDriveLogoDataUri_();

  return template.evaluate()
    .setTitle('Workshop Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDriveLogoDataUri_() {
  const logoFileId = '1091teRqeYeBZ-v601l15lCl7iPjyjOu5';
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob();
  const contentType = logoBlob.getContentType() || 'image/png';
  const base64Data = Utilities.base64Encode(logoBlob.getBytes());
  return `data:${contentType};base64,${base64Data}`;
}

// THE GLUE: This allows us to inject CSS and JS files into Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
