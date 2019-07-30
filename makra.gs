var source = SpreadsheetApp.getActiveSpreadsheet();
var firstTerritory = 1;
var lastTerritory = 142;

function goTo() {
  source.getSheetByName('Lista terenów').activate();
};

function fasterSheetActivation() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Podaj numer");
//  if (response.getSelectedButton() == ui.Button.OK) {
    var textFrom = response.getResponseText();
    var textNumber = parseFloat(textFrom);
  source.getSheetByName('Teren nr ' + textNumber).activate();
//  };
};


function set_territory_type_condition() {
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.getRange("E6").setFormula('=if(COUNTIF(A8:A; "<>") = COUNTIF(B8:B; "<>"); if(filter(C8:C; B8:B = MAX(B8:B))=""; "grupowy"; "indywidualny"); if(filter(C8:C; A8:A = MAX(A8:A))<>""; "grupowy"; if(filter(D8:D; A8:A = MAX(A8:A))<>""; "indywidualny"; if(filter(C8:C; B8:B = MAX(B8:B))=""; "grupowy"; "indywidualny"))))');
  };
  
};

function set_territory_owner_condition() {
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.getRange("E5").setFormula('=IF(E3=TRUE;IF(FILTER(C8:C;A8:A=E4)="";FILTER(D8:D;A8:A=E4);FILTER(C8:C;A8:A=E4));"")').setHorizontalAlignment("center").setVerticalAlignment("middle");
  };
 
}

function set_cond_format_rules() {
  var ui = SpreadsheetApp.getUi();
  
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.clearConditionalFormatRules();
    var rules = territorySheet.getConditionalFormatRules();
    
    var terr_type_range = territorySheet.getRange("E6");
    var header_group_range = territorySheet.getRange("C7");
    var header_publisher_range = territorySheet.getRange("D7");
    var group_entries_range = territorySheet.getRange("C8:C");
    var publishers_entries_range = territorySheet.getRange("D8:D");
    
    terr_type_range.setFontFamily("Arial").setHorizontalAlignment("center").setFontSize(10);
    
    // terr type jest grupowy
     var terr_type_gr_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="grupowy"'])
    .setBackground("#F3F3F3")
    .setFontColor("#000000")
    .setRanges([terr_type_range])
    .build();     

    // terr type jest indywidualny
     var terr_type_ind_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="indywidualny"'])
    .setBackground("#F3F3F3")
    .setFontColor("#000000")
    .setRanges([terr_type_range])
    .build();  
    
     // C6
     var header_group_range_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="grupowy"'])
    .setBold(true)
    .setBackground("#1F114B")
    .setRanges([header_group_range])
    .build(); 
    
    // D6
     var header_publisher_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="indywidualny"'])
    .setBold(true)
    .setBackground("#0C343D")
    .setRanges([header_publisher_range])
    .build(); 
    
    // C7:C kiedy nie jest grupowy
    var group_entries_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="indywidualny"'])
    .setFontColor("#434343")
    .setRanges([group_entries_range])
    .build();

    // C7:C kiedy jest grupowy
    var group_entries_rule2 = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="grupowy"'])
    .setBackground("#EFEFEF")
    .setBold(true)
    .setRanges([group_entries_range])
    .build();    
    
    // D7:D kiedy nie jest indywidualny
    var publishers_entries_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="grupowy"'])
    .setFontColor("#434343")
    .setRanges([publishers_entries_range])
    .build();
    
    // D7:D kiedy jest indywidualny
    var publishers_entries_rule2 = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$6="indywidualny"'])
    .setBackground("#ECF5FC")
    .setBold(true)
    .setRanges([publishers_entries_range])
    .build();
    
    rules.push(terr_type_gr_rule);
    rules.push(terr_type_ind_rule);
    rules.push(header_group_range_rule);
    rules.push(header_publisher_rule);
    rules.push(group_entries_rule);
    rules.push(group_entries_rule2);
    rules.push(publishers_entries_rule);
    rules.push(publishers_entries_rule2);
    territorySheet.setConditionalFormatRules(rules);
  };

};

function get_territory_type_from_terrtory_card() {
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!$E$5"]);
  };
  var range = sheet.getRange("M" + (firstTerritory+1) + ":M" + (lastTerritory+1));
  range.setValues(lista);
}

//ustawia regułu zliczania opracowań terenu od podanej daty w arkuszu Lista zamian terenów pole G2
function set_last_territory_changes_counter() {
  var formulas = [];
  var row2 = [];
  var row3 = [];
  row2.push('=COUNT(filter(B8:B; A8:A < \'Lista zamian terenów\'!G2))');
  row2.push('=IF(A2=1;"raz opracowano teren przed zmianą " & TEXT(\'Lista zamian terenów\'!G2; "yyyy-mm-dd"); "razy opracowano teren przed zmianą " & TEXT(\'Lista zamian terenów\'!G2; "yyyy-mm-dd"))');
  formulas.push(row2);
  row3.push('=COUNT(filter(B8:B; A8:A > \'Lista zamian terenów\'!G2; B8:B > \'Lista zamian terenów\'!G2))');
  row3.push('=IF(A3=1;"raz opracowano teren od zmiany " & TEXT(\'Lista zamian terenów\'!G2; "yyyy-mm-dd"); "razy opracowano teren od zmiany " & TEXT(\'Lista zamian terenów\'!G2; "yyyy-mm-dd"))');
  formulas.push(row3);
  for (var number = 1; number <= 142; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.getRange("A3:B4").setFormulas(formulas);
};
}

function set_territory_assigned() {
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.getRange("E3").setFormula('=IF(COUNTA(A8:A)>COUNTA(B8:B);TRUE;FALSE)');
  };
}

function extend_territory_card() {

  for (var number = 12; number <= 12; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    if (territorySheet.getMaxColumns() < 7){
       territorySheet.insertColumnAfter(5);
       territorySheet.insertColumnAfter(6);
       territorySheet.insertRowAfter(1)
       }
    var value_tmp = territorySheet.getRange("D1").getValue();
    var formula_tmp = territorySheet.getRange("D1").getFormula();
    territorySheet.getRange("A:G").setFontFamily("Roboto").setVerticalAlignment("middle");
    territorySheet.getRange("A2:G2").merge().setBackground("#F3F3F3");
    territorySheet.getRange("A2").setFormula("='Lista terenów'!B" + (number + 1)).setFontSize(11).setBackground("#F3F3F3").setFontColor("#32255B");
    territorySheet.getRange("A2").setFontStyle("italic").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setFontWeight("normal").setFontSize(11);
    territorySheet.getRange("A3").setFormula("=COUNT(filter(B8:B; A8:A < 'Lista zamian terenów'!G2))");
    territorySheet.getRange("A4").setFormula("=COUNT(filter(B8:B; A8:A > 'Lista zamian terenów'!G2; B8:B > 'Lista zamian terenów'!G2))");
    territorySheet.getRange("A5").setFormula("=MAX(B8:B)");
    territorySheet.getRange("E1").setValue("STATUS").setBackground("#741B47").setFontSize(11);
    territorySheet.getRange("E7").setValue("Grupowy");
    territorySheet.getRange("E5").clearContent().clear();
    territorySheet.getRange("E3").setBackground("#F1C232").setFontColor("#000000");   
    territorySheet.getRange("E4").setBackground("#FFD966").setFontColor("#000000");   
    territorySheet.getRange("E5").setBackground("#FFE599").setFontColor("#000000");  
    territorySheet.getRange("F7").setValue("Kampania");
    territorySheet.getRange("G7").setValue("Uwagi");
    territorySheet.getRange("A1:F1").merge().setValue("Teren nr "+ number).setBackground("#1C4587");
    territorySheet.getRange("G1").setValue(value_tmp).setFormula(formula_tmp).setBackground("#1C4587");
    territorySheet.getRange("F3:G3").setBackground("#6D9EEB").merge().setValue("ADNOTACJE").setFontColor("#32255B");
    territorySheet.getRange("F4:G6").setBackground("#F3F3F3").merge();
    territorySheet.getRange("A7:G7").setBackground("#0B5394");
    //territorySheet.getRange("E6").setBackground("#073763");
    territorySheet.getRange("A3:D3").setBackground("#6FA8DC");
    territorySheet.getRange("A4:D4").setBackground("#9FC5E8");
    territorySheet.getRange("A5:D5").setBackground("#CFE2F3");
    territorySheet.getRange("A6:D6").setBackground("#ECF5FC");
    territorySheet.getRange("A3:A6").setHorizontalAlignment("center");
    territorySheet.setRowHeight(2, 37);
    territorySheet.setRowHeights(3, 4, 30);
    territorySheet.setColumnWidths(1, 6, 110);
    };
};


