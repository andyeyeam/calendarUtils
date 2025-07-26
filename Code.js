function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Andy's Calendar Utilities")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onAvailabilityFinderClick() {
  Logger.log('Availability Finder button clicked');
  // Add your availability finder logic here
  return 'Availability Finder functionality will be implemented here xx';
}

function onSkipLevelManagerClick() {
  Logger.log('Skip Level Manager button clicked');
  // Add your skip level manager logic here
  return 'Skip Level Manager functionality will be implemented here';
}