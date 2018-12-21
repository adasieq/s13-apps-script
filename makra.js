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
    territorySheet.getRange("E5").setFormula('=if(E2=FALSE; if(filter(C7:C; B7:B = MAX(B7:B))=""; "grupowy"; "indywidualny"); if(filter(C7:C; B7:B = MAX(B7:B))=""; "grupowy"; "indywidualny"))');
  };
  
};

function set_cond_format_rules() {
  var ui = SpreadsheetApp.getUi();
  
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    
    var territorySheet = source.getSheetByName("Teren nr " + number);
    territorySheet.clearConditionalFormatRules();
    var rules = territorySheet.getConditionalFormatRules();
    
    var terr_type_range = territorySheet.getRange("E5");
    var header_group_range = territorySheet.getRange("C6");
    var header_publisher_range = territorySheet.getRange("D6");
    var group_entries_range = territorySheet.getRange("C7:C");
    var publishers_entries_range = territorySheet.getRange("D7:D");
    
    terr_type_range.setFontFamily("Arial").setHorizontalAlignment("center").setFontSize(10);
    territorySheet.autoResizeColumn(4);
    
    // terr type jest grupowy
     var terr_type_gr_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="grupowy"'])
    .setBackground("#E0E9FB")
    .setFontColor("#000000")
    .setRanges([terr_type_range])
    .build();     

    // terr type jest indywidualny
     var terr_type_ind_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="indywidualny"'])
    .setBackground("#e4dcf5")
    .setFontColor("#000000")
    .setRanges([terr_type_range])
    .build();  
    
     // C6
     var header_group_range_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="grupowy"'])
    .setBold(true)
    .setBackground("#1F114B")
    .setRanges([header_group_range])
    .build(); 
    
    // D6
     var header_publisher_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="indywidualny"'])
    .setBold(true)
    .setBackground("#4B102E")
    .setRanges([header_publisher_range])
    .build(); 
    
    // C7:C kiedy nie jest grupowy
    var group_entries_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="indywidualny"'])
    .setFontColor("#434343")
    .setRanges([group_entries_range])
    .build();

    // C7:C kiedy jest grupowy
    var group_entries_rule2 = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="grupowy"'])
    .setBackground("#E0E9FB")
    .setBold(true)
    .setRanges([group_entries_range])
    .build();    
    
    // D7:D kiedy nie jest indywidualny
    var publishers_entries_rule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="grupowy"'])
    .setFontColor("#434343")
    .setRanges([publishers_entries_range])
    .build();
    
    // D7:D kiedy jest indywidualny
    var publishers_entries_rule2 = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=$E$5="indywidualny"'])
    .setBackground("#e4dcf5")
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
