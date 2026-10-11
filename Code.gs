var ss = SpreadsheetApp.getActiveSpreadsheet()
var form = ss.getSheetByName('Insert เก็บประวัติการซื้อขายของลูกค้า (Sale)')
var data = ss.getSheetByName('ประวัติการซื้อขายของลูกค้า All')

function Submit() {
  var values = [[form.getRange("E6").getValue(),
  form.getRange("E8").getValue(),
  form.getRange("E10").getValue(),
  form.getRange("E12").getValue(),
  form.getRange("E14").getValue(),
  form.getRange("E16").getValue(),
  form.getRange("E20").getValue(),
  form.getRange("G20").getValue(),
  form.getRange("E22").getValue(),
  form.getRange("G22").getValue(),
  form.getRange("E24").getValue(),
  form.getRange("G24").getValue(),
  form.getRange("E26").getValue(),
  form.getRange("G26").getValue(),
  form.getRange("E28").getValue(),
  form.getRange("G28").getValue(),
  form.getRange("E30").getValue(),
  form.getRange("G30").getValue(),
  form.getRange("E32").getValue(),
  form.getRange("G32").getValue(),
  form.getRange("E34").getValue(),
  form.getRange("G34").getValue(),
  form.getRange("E36").getValue(),
  form.getRange("G36").getValue(),
  form.getRange("E38").getValue(),
  form.getRange("G38").getValue(),
  form.getRange("E40").getValue(),
  form.getRange("E42").getValue(),
  form.getRange("E44").getValue()]]
  data.getRange(data.getLastRow() + 1, 1, 1, 29).setValues(values)
  Clear()
}

function Clear() {
  var clear = ["E3", "E6", "E8", "E10", "E12", "E14", "E16", "E20", "G20", "E22", "G22", "E24", "G24", "E26","G26", "E28", "G28",  "E30", "G30", "E32", "G32", "E34", "G34", "E36", "G36", "E38", "G38", "E40", "E42", "E44"]
  for (var i = 0; i < clear.length; i++) {
    form.getRange(clear[i]).clearContent()
  }
}


var searchID = 0
function Search() {
  var str = form.getRange("E3").getValue()
  var values = ss.getSheetByName('ประวัติการซื้อขายของลูกค้า All').getDataRange().getValues()
  for (var i = 0; i < values.length; i++) {
    var row = values[i]
    if (row[searchID] == str) {
      form.getRange("E6").setValue(row[0])
      form.getRange("E8").setValue(row[1])
      form.getRange("E10").setValue(row[2])
      form.getRange("E12").setValue(row[3])
      form.getRange("E14").setValue(row[4])
      form.getRange("E16").setValue(row[5])
      form.getRange("E20").setValue(row[6])
      form.getRange("G20").setValue(row[7])
      form.getRange("E22").setValue(row[8])
      form.getRange("G22").setValue(row[9])
      form.getRange("E24").setValue(row[10])
      form.getRange("G24").setValue(row[11])
      form.getRange("E26").setValue(row[12])
      form.getRange("G26").setValue(row[13])
      form.getRange("E28").setValue(row[14])
      form.getRange("G28").setValue(row[15])
      form.getRange("E30").setValue(row[16])
      form.getRange("G30").setValue(row[17])
      form.getRange("E32").setValue(row[18])
      form.getRange("G32").setValue(row[19])
      form.getRange("E34").setValue(row[20])
      form.getRange("G34").setValue(row[21])
      form.getRange("E36").setValue(row[22])
      form.getRange("G36").setValue(row[23])
      form.getRange("E38").setValue(row[24])
      form.getRange("G38").setValue(row[25])
      form.getRange("E40").setValue(row[26])
      form.getRange("E42").setValue(row[27])
      form.getRange("E44").setValue(row[28])
    }
  }
}

function Update() {
  var str = form.getRange("E3").getValue()
  var values = ss.getSheetByName('ประวัติการซื้อขายของลูกค้า All').getDataRange().getValues()
  for (var i = 0; i < values.length; i++) {
    var row = values[i]
    if (row[searchID] == str) {
      var int = i + 1
      var values1 = [[form.getRange("E6").getValue(),
      form.getRange("E8").getValue(),
      form.getRange("E10").getValue(),
      form.getRange("E12").getValue(),
      form.getRange("E14").getValue(),
      form.getRange("E16").getValue(),
      form.getRange("E20").getValue(),
      form.getRange("G20").getValue(),
      form.getRange("E22").getValue(),
      form.getRange("G22").getValue(),
      form.getRange("E24").getValue(),
      form.getRange("G24").getValue(),
      form.getRange("E26").getValue(),
      form.getRange("G26").getValue(),
      form.getRange("E28").getValue(),
      form.getRange("G28").getValue(),
      form.getRange("E30").getValue(),
      form.getRange("G30").getValue(),
      form.getRange("E32").getValue(),
      form.getRange("G32").getValue(),
      form.getRange("E34").getValue(),
      form.getRange("G34").getValue(),
      form.getRange("E36").getValue(),
      form.getRange("G36").getValue(),
      form.getRange("E38").getValue(),
      form.getRange("G38").getValue(),
      form.getRange("E40").getValue(),
      form.getRange("E42").getValue(),
      form.getRange("E44").getValue()]]
      data.getRange(int, 1, 1, 29).setValues(values1)
      SpreadsheetApp.getUi().alert('Updated')
    }
  }
}

function Delete() {
  var ui = SpreadsheetApp.getUi()
  var response = ui.alert('Are you sure you want to delete this data? This Function may delete this data forever.', ui.ButtonSet.YES_NO)
  if (response == ui.Button.YES) {
    var str = form.getRange("E3").getValue()
    var values = ss.getSheetByName('ประวัติการซื้อขายของลูกค้า All').getDataRange().getValues()
    for (var i = 0; i < values.length; i++) {
      var row = values[i]
      if (row[searchID] == str) {
        var int = i + 1
        data.deleteRow(int)
        Clear()
      }
    }
  }
}

function SaveData() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:A').activate();
  spreadsheet.duplicateActiveSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Graph(จังหวัด) All'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Graph(ประเภทลูกค้า) All'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('รายการ(1/2) All'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('รายการ(2/2) All'), true);
  spreadsheet.duplicateActiveSheet();
}

function ClearForm() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:AB1000000').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
}

var searchID = 0
function BringUp() {
  var str = form.getRange("E3").getValue()
  var values = ss.getSheetByName('Primary Data/Exps.').getDataRange().getValues()
  for (var i = 0; i < values.length; i++) {
    var row = values[i]
    if (row[searchID] == str) {
      form.getRange("E6").setValue(row[0])
      form.getRange("E8").setValue(row[1])
      form.getRange("E10").setValue(row[2])
      form.getRange("E12").setValue(row[3])
      form.getRange("E14").setValue(row[4])
      form.getRange("E16").setValue(row[5])
      form.getRange("E42").setValue(row[6])
      form.getRange("E44").setValue(row[7])
    }
  }
}
