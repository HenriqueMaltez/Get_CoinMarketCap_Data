function add_menu() {

  //Adding a menu to the table with the functions we created
  let ui = SpreadsheetApp.getUi();
  
  ui.createMenu('My Functions')
    .addItem('Update Tokens Data', 'getdata')
    .addItem('Update Meta Data', 'getmetadata')
    .addToUi();
}
