function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.logoDataUri = getDriveLogoDataUri_();

  return template.evaluate()
    .setTitle('Workshop Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDriveLogoDataUri_() {
  try {
    const logoFileId = '17w396Wt457eccDZrqiEkK1Y8l6YtAYIu';
    const logoBlob = DriveApp.getFileById(logoFileId).getBlob();
    const contentType = logoBlob.getContentType() || 'image/png';
    const base64Data = Utilities.base64Encode(logoBlob.getBytes());
    return `data:${contentType};base64,${base64Data}`;
  } catch (e) {
    return '';
  }
}

// THE GLUE: This allows us to inject CSS and JS files into Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
