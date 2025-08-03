function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Andy\'s Calendar Utilities')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getIndexPage() {
  return HtmlService.createHtmlOutputFromFile('index').getContent();
}