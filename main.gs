function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Workshop Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// THE GLUE: This allows us to inject CSS and JS files into Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
