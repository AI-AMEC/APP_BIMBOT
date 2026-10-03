// ==========================================
// BIMBOT v41.0 - 99_WebApp
// Front-end Deployment Entry Point & Template Engine
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AMEC BIMBOT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Splits the HTML into modular components. Evaluated as a template so
// views can include sub-views (e.g. View_ProjectDetails -> View_Publish).
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
