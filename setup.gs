// function setup accepts and stores the Consumer Key, Consumer Secret, Project Key, firstDate, and list of Data Elements
function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Setup Fitbit Download");
    app.setStyleAttribute("padding", "10px");

    var consumerKeyLabel = app.createLabel("Fitbit OAuth 2.0 Client ID:*");
    var consumerKey = app.createTextBox();
    consumerKey.setName("consumerKey");
    consumerKey.setWidth("100%");
    consumerKey.setText(getConsumerKey());

    var consumerSecretLabel = app.createLabel("Fitbit OAuth Consumer Secret:*");
    var consumerSecret = app.createTextBox();
    consumerSecret.setName("consumerSecret");
    consumerSecret.setWidth("100%");
    consumerSecret.setText(getConsumerSecret());

    var projectKeyTitleLabel = app.createLabel("Project key: ");
    var projectKeyLabel = app.createLabel(ScriptApp.getProjectKey());

    var firstDate = app.createTextBox().setId("firstDate").setName("firstDate");
    firstDate.setName("firstDate");
    firstDate.setWidth("100%");
    firstDate.setText(getFirstDate());

    // create the save handler and button
    var saveHandler = app.createServerClickHandler("saveSetup");
    var saveButton = app.createButton("Save Setup", saveHandler);

    // put the controls in a grid
    var listPanel = app.createGrid(8, 3);
    listPanel.setWidget(1, 0, consumerKeyLabel);
    listPanel.setWidget(1, 1, consumerKey);
    listPanel.setWidget(2, 0, consumerSecretLabel);
    listPanel.setWidget(2, 1, consumerSecret);
    listPanel.setWidget(3, 0, app.createLabel(" * (obtain these at dev.fitbit.com, use OAuth2.0)"));
    listPanel.setWidget(4, 0, projectKeyTitleLabel);
    listPanel.setWidget(4, 1, projectKeyLabel);
    listPanel.setWidget(5, 0, app.createLabel("Download day (yyyy-mm-dd or 'today')"));
    listPanel.setWidget(5, 1, firstDate);
    listPanel.setWidget(7, 0, app.createLabel("Very long intervals will not work; exceed Fitbit rate limit and/or function will timeout"));

    // Ensure that all controls in the grid are handled
    saveHandler.addCallbackElement(listPanel);
    // Build a FlowPanel, adding the grid and the save button
    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(listPanel);
    dialogPanel.add(saveButton);
    app.add(dialogPanel);
    doc.show(app);
}
