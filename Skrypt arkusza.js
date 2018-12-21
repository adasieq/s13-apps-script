/*
####################################################################################
####################################################################################
####################################################################################
   Tu możesz edytować ilość terenów i podział terenów na grupy.
   Poniżej możesz ustawić numer ostatniego ternu (lastTerritory).
   Musi być on zgodny z ilością założonych kart terenów w arkuszu.
####################################################################################
####################################################################################
*/

var firstTerritory = 1;
var lastTerritory = 142;

//##################################################################################
// Tu odbywa się dzielenie terenów na 3 części. Powinno to wyglądać tak: 
// var groupX = [1, 2, 3, 4, 5];
// gdzie X to numer grupy
//##################################################################################

var group1 = [1,  2,  9,  10,  15,  16,  17,  18,  19,  29,  34,  35,  38,  41,  50,  51,  52,  54,  56,  58,  60,  62,  65,  68,  71,  74,  77,  80,  83,  86,  90,  92,  95,  98,  101,  104,  107,  110,  113,  114,  115,  122,  127,  128,  131,  140,  141,  142];
var group2 = [3,  4,  8,  11,  14,  20,  22,  28,  30,  31,  33,  36,  37,  42,  47,  48,  49,  55,  57,  61,  63,  66,  69,  72,  75,  78,  81,  84,  87,  88,  93,  96,  99,  102,  105,  108,  111,  116,  117,  118,  123,  129,  130,  132,  136,  137,  138];
var group3 = [5,  6,  7,  12,  13,  21,  23,  24,  25,  26,  27,  32,  39,  40,  43,  44,  45,  46,  53,  59,  64,  67,  70,  73,  76,  79,  82,  85,  89,  91,  94,  97,  100,  103,  106,  109,  112,  119,  120,  121,  124,  125,  126,  133,  134,  135,  139];
  
/*
####################################################################################
####################################################################################
    To by było na tyle.
####################################################################################
####################################################################################
####################################################################################
*/

var source = SpreadsheetApp.getActiveSpreadsheet();

function goTo() {
  source.getSheetByName('Lista terenów').activate();
};

function wz() {
  source.getSheetByName('WZ*').activate();
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

function copyEntriesWZmanually() {
  var ui = SpreadsheetApp.getUi();
  var areYouSure = ui.alert("Czy jesteś pewien, że chcesz zaktualizować WZ*? Operacja ta jest niezbędna tylko do korzystania z aktualnych statystyk. Arkusze te są oznaczone gwiazdką.", ui.ButtonSet.YES_NO);
  if (areYouSure == ui.Button.YES) {
    copyEntriesWZ()
  }
}

function timeStamp() {
  source.getSheetByName("WZ*").getRange(2, 1).setValue(new Date().toLocaleString());
}

function successfulExecution() {
  //Potwierdzenie
  source.toast('Funkcja została pomyślnie wykonana', 'Udało się!', 10);
};

function onOpen() {
  //Stworzenie dodatkowego przycisku na pasku menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Funkcje skryptowe')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('WZ*...')
              .addItem('Wyczyść WZ*', 'deleteEntriesWZ')
              .addItem('Aktualizuj WZ*', 'copyEntriesWZmanually')
              .addSeparator()
              .addItem('Sortuj rosnąco wg daty przydzielenia', 'sortWZAssignDateAscending')
              .addItem('Sortuj malejąco wg daty przydzielenia', 'sortWZAssignDateDescending')
              .addSeparator()
              .addItem('Sortuj rosnąco wg numeru i daty przydzielenia', 'sortWZTerrAscendingAssignDateAscending')
              .addItem('Sortuj rosnąco wg numeru i malejąco wg daty przydzielenia', 'sortWZTerrAscendingAssignDateDescending')
              .addSeparator()
              .addItem('Sortuj rosnąco wg daty zdania', 'sortWZReturnDateAscending')
              .addItem('Sortuj malejąco wg daty zdania', 'sortWZReturnDateDescending')
              .addSeparator()
              .addItem('Sortuj rosnąco wg numeru i daty zdania', 'sortWZTerrAscendingReturnDateAscending')
              .addItem('Sortuj rosnąco wg numeru i malejąco wg daty zdania', 'sortWZTerrAscendingReturnDateDescending')
             )
  .addItem('Zmiana terenów osobistych i grupowych', 'changeAssignments')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Dodawanie i usuwanie arkuszy')
              .addItem('Dodaj arkusze ze wzorca', 'copySheets')
              .addItem('Usuń dodane arkusze', 'deleteSheets'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Napraw...')
              .addItem('Napraw spacje', 'fixSpaces')
              .addItem('Napraw błedne nazwiska', 'fixWrongNames')
              .addItem('Napraw linki na liście', 'makeLinks')
              .addItem('Napraw sekcję "Przydzielony"', 'Assigned')
              .addItem('Napraw sekcję "Opracowano"', 'recentlyVisited'))
  .addToUi();
  ui.createMenu('Przejdź do...')
  .addItem('Narzędzia', 'show_territory_cards_tools')
  .addItem('Boczna tabela terenów', 'territoryTables')
  .addItem('Lista terenów', 'goTo')
  .addItem('*WZ', 'wz')
  .addItem('Terenu nr...', 'fasterSheetActivation')
  .addToUi();
};

function copySheets() {
  //Zapytanie o początek zakresu do stworzenia nowych arkuszy na podstawie wzorca
  var ui = SpreadsheetApp.getUi();
  var response1 = ui.prompt("Od jakiego numeru rozpocząć generowanie?");
  if (response1.getSelectedButton() == ui.Button.OK) {
    var textFrom = response1.getResponseText();
    var from = parseFloat(textFrom);
    //Zapytanie o koniec tego zakresu
    var ui = SpreadsheetApp.getUi();    
    var response2 = ui.prompt("Na jakim numerze skończyć generowanie?");
    if (response2.getSelectedButton() == ui.Button.OK) {
      var textTo = response2.getResponseText();
      var to = parseFloat(textTo);
      //Potwierdzenie zmian
      var ui = SpreadsheetApp.getUi();
      var howMany = to - from + 1;
      var response3 = ui.alert("Czy jesteś pewien, że chcesz wygenerować nowe arkusze w ilości: " + howMany + "?", ui.ButtonSet.YES_NO);
      if (response3 == ui.Button.YES) {
        //Stworzenie nowych arkuszy i skopiowanie nr terenu
        for (var number = from; number <= to; number = number + 1) {
          var motherSheet = source.getSheetByName("Wzorzec");
          var motherSheetName = motherSheet.getName();
          var destination = source;
          motherSheet.copyTo(destination);
          var childSheet = source.getSheetByName("Kopia arkusza" + " " + motherSheetName);
          childSheet.setName("Teren nr" + " " + number);
          var insideNumber = childSheet.setActiveSelection("C1");
          insideNumber.setValue(number);
        };
      };
    };
  };
  successfulExecution();
};

function deleteSheets() {
  //Zapytanie o początek zakresu arkuszy do usunięcia
  var ui = SpreadsheetApp.getUi();
  var response1 = ui.prompt("Od jakiego numeru rozpocząć USUWANIE?");
  if (response1.getSelectedButton() == ui.Button.OK) {
    var textFrom = response1.getResponseText();
    var from = parseFloat(textFrom);
    //Zapytanie o koniec tego zakresu
    var ui = SpreadsheetApp.getUi();
    var response2 = ui.prompt("Na jakim numerze skończyć USUWANIE?");
    if (response2.getSelectedButton() == ui.Button.OK) {
      var textTo = response2.getResponseText();
      var to = parseFloat(textTo);
      //Potwiedzenie wykonania
      var ui = SpreadsheetApp.getUi();
      var howMany = to - from + 1;
      var response3 = ui.alert("Czy jesteś pewien, że chcesz USUNĄĆ arkusze od " + from + " do " + to + " (łącznie " + howMany + ")?", ui.ButtonSet.YES_NO);
      if (response3 == ui.Button.YES) {
        var responseDel = ui.prompt('Jeżeli jesteś pewien, wpisz "Usuwam" poniżej');
        if (responseDel.getSelectedButton() == ui.Button.OK) {
          var textDel = responseDel.getResponseText();
          if (textDel === "Usuwam") {
            //Usuwanie arkuszy
            for (var number = from; number <= to; number = number + 1) {
              var sheet = source.getSheetByName("Teren nr " + number);
              source.toast("Usuwanie arkusza Teren nr " + number);
              source.deleteSheet(sheet);
            };
          } else {
            source.toast('Funkcja została anulowana', 'Błąd', 10);
            return "Funkcja została anulowana"
          };
        };
      };
    };
  };
  successfulExecution();
};

function makeLinks() {
  //Wykonanie linków
  var list = [];
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var sheet = source.getSheetByName("Teren nr " + number);
    var ID = sheet.getSheetId();
    var value = (['=HYPERLINK("#gid=' + ID + '";"' + number + '")']);
    list.push(value);
  };
  source.getSheetByName("Lista terenów")
  .getRange("A" + (firstTerritory + 1) + ":A" + (lastTerritory + 1))
  .setValues(list);
  successfulExecution();
};

function deleteEntriesWZ() {
  //Usunięcie starych wartości z WZ
  var WZ = source.getSheetByName("WZ*");
  var WZRange = WZ.getRange("A6:G");
  WZRange.clearContent();
}

function copyEntriesWZ() {
  deleteEntriesWZ();
  var WZ = source.getSheetByName("WZ*");
  source.toast('Rozpoczęto aktualizację WZ*.  Trochę to potrwa...', 'Uruchomiono skrypt', 60);
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var WZLastRow = WZ.getLastRow();
    var motherSheet = source.getSheetByName("Teren nr " + number);
    var motherLastRowNumber = motherSheet.getLastRow();
    if (motherLastRowNumber >= 7) {
      var WZrange = WZ.getRange((WZLastRow+1), 2,(motherLastRowNumber-6), 5);
      source.getUrl()
      var sheet = source.getSheetByName("Teren nr " + number);
      var ID = sheet.getSheetId();
      var territory_link = (['=HYPERLINK("'+ source.getUrl() +'#gid=' + ID + '";"' + number + '")']);  
      
      WZ.getRange((WZLastRow+1), 1, (motherLastRowNumber-6)).setValue(territory_link);
      motherSheet.getRange(7, 1, (motherLastRowNumber-6), 5).copyTo(WZrange, {contentsOnly:true});
      WZ.getRange((WZLastRow+1), 7, (motherLastRowNumber-6)).setValue(getTerritoryGroup(number));
      WZ.getRange((WZLastRow+1), 7, (motherLastRowNumber-6)).setFormula('=if(AND(D'+(WZLastRow+1)+'="";E'+(WZLastRow+1)+'<>"");"indywidualny"; "grupowy")');
    };
  };
  timeStamp();
  successfulExecution();
}

function sortWZAssignDateAscending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort(2);
  successfulExecution()
};

function sortWZAssignDateDescending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort({column: 2, ascending: false});
  successfulExecution()
};

function sortWZTerrAscendingAssignDateAscending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort([{column: 1}, {column: 2}]);
  successfulExecution()
};

function sortWZTerrAscendingAssignDateDescending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort([{column: 1}, {column: 2, ascending: false}]);
  successfulExecution()
};

function sortWZReturnDateAscending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort(3);
  successfulExecution()
};

function sortWZReturnDateDescending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort({column: 3, ascending: false});
  successfulExecution()
};

function sortWZTerrAscendingReturnDateAscending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort([{column: 1}, {column: 3}]);
  successfulExecution()
};

function sortWZTerrAscendingReturnDateDescending() {
  var WZ = source.getSheetByName("WZ*");
  var WZLastRow = WZ.getLastRow();
  var WZRange = WZ.getRange(6, 1, WZLastRow, 5);
  WZRange.sort([{column: 1}, {column: 3, ascending: false}]);
  successfulExecution()
};

function WZnote() {
  //Przypomnienie w postaci toasta o konieczności aktualizacji WZ po wprowadzeniu danych (patrz Wyzwalacze).
  source.toast('Wprowadziłeś zmiany. Pamiętaj o zaktualizowaniu WZ*. WZ* aktualizuje się automatycznie 1 raz dziennie w godzinach 3:00 - 4:00.', "Wprowadzono zmiany", 25);
};

function newListMaker(motherList, childList) {
  
  for (var i = 0; i < motherList.length; i++) {
    childList.push(motherList[i]);
  };
};

function changeAssignments() {
 //Wybór grupy do terenów osobistych
  var ui = SpreadsheetApp.getUi();
  var personalGroup = ui.prompt("Która część terenu ma zostać terenami osobistymi? Wpisz 1, 2 lub 3.");
  if (personalGroup.getSelectedButton() == ui.Button.OK) {
    var text = personalGroup.getResponseText();
    var personalGroup = parseFloat(text);
    if (personalGroup === 1) {
      var personalGroup = [];
      var groupGroup = [];
      newListMaker(group1, personalGroup);
      newListMaker(group2, groupGroup);
      newListMaker(group3, groupGroup);
    } else if (personalGroup === 2) {
      var personalGroup = [];
      var groupGroup = [];
      newListMaker(group2, personalGroup);
      newListMaker(group1, groupGroup);
      newListMaker(group3, groupGroup);
    } else if (personalGroup === 3) {
      var personalGroup = [];
      var groupGroup = [];
      newListMaker(group3, personalGroup);
      newListMaker(group1, groupGroup);
      newListMaker(group2, groupGroup);
    } else {
      var ui = SpreadsheetApp.getUi();
      personalGroup = null;
      groupGroup = null;
      var response = ui.alert("Nie podano prawidłowej wartości", ui.ButtonSet.OK);
      return "Nie podano prawidłowej wartości";
    };
    var text = [];
    var background = [];
    var fontColor = [];
    var groupTypeNumber = [];
    for (var i = firstTerritory; i <= lastTerritory; i++) {
      if (personalGroup.indexOf(i) != (-1)) {
        text.push(["indywidualny"]);
        background.push(["#515ec4"]);
        fontColor.push(["#e1eafc"]);
      } else if (groupGroup.indexOf(i) != (-1)) {
        text.push(["grupowy"]);
        background.push(["#b1c2f6"]);
        fontColor.push(["#293187"]);
      };
  
      if (group1.indexOf(i) != (-1)) {
        groupTypeNumber.push(['1']);
      } else if (group2.indexOf(i) != (-1)) {
        groupTypeNumber.push(['2']);
      } else if (group3.indexOf(i) != (-1))  {
        groupTypeNumber.push(['3']);
      };
      
    };
    var sheet = source.getSheetByName("Lista terenów");
    var range = sheet.getRange("J" + (firstTerritory+1) + ":J" + (lastTerritory+1));
    var range2 = sheet.getRange("K" + (firstTerritory+1) + ":K" + (lastTerritory+1));
    range.setValues(text);
    range.setBackgrounds(background);
    range.setFontColors(fontColor);
    range2.setValues(groupTypeNumber);
  };
  successfulExecution();
};


function getTerritoryGroup(territoryNumber) {
  
  var result = 0;
  
      if (group1.indexOf(territoryNumber) != (-1)) {
        result = 1;
      } else if (group2.indexOf(territoryNumber) != (-1)) {
        result = 2;
      } else if (group3.indexOf(territoryNumber) != (-1))  {
        result = 3;
      };
  
  return result;
}


//Aktualizacja sekcji "Opracowano" (całość)
function recentlyVisited() {
  //w tym roku
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!A3"]);
  };
  var range = sheet.getRange("H" + (firstTerritory+1) + ":H" + (lastTerritory+1));
  range.setValues(lista);
  //poprzednim
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!A2"]);
  };
  var range = sheet.getRange("I" + (firstTerritory+1) + ":I" + (lastTerritory+1));
  range.setValues(lista);
  //Opr ostatnio
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!A5"]);
  };
  var range = sheet.getRange("G" + (firstTerritory+1) +":G" + (lastTerritory+1));
  range.setValues(lista);
  successfulExecution();
};

//Aktualizacja sekcji "Przydzielono" (całość)
function Assigned() {
  //Przydzielony
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!$E$2"]);
  };
  var range = sheet.getRange("C" + (firstTerritory+1) + ":C" + (lastTerritory+1));
  range.setValues(lista);
  //kiedy?
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!$E$3"]);
  };
  var range = sheet.getRange("D" + (firstTerritory+1) + ":D" + (lastTerritory+1));
  range.setValues(lista);
  //jak dawno?
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = (firstTerritory+1); i <= (lastTerritory+1); i++) { 
    lista.push(["=IF(C" +i + '=TRUE;TODAY()-D' +i +';"")']);
  };
  var range = sheet.getRange("E" + (firstTerritory+1) + ":E" + (lastTerritory+1));
  range.setValues(lista);
  //komu?
  var sheet = source.getSheetByName("Lista terenów");
  var lista = [];
  for (var i = firstTerritory; i <= lastTerritory; i++) {
    lista.push(["='Teren nr " + i + "'!$E$4"]);
  };
  var range = sheet.getRange("F" + (firstTerritory+1) + ":F" + (lastTerritory+1));
  range.setValues(lista);
  successfulExecution();
}

function fixSpaces() {
  var ui = SpreadsheetApp.getUi();
  var first_territroy_range_data_row = 7;
  var publisher_name_column = 4;
  var fix_list = [];
  var re_publisher_spaces = new RegExp(/\w+\.\s+\w+/);
  
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    var territory_range_data = territorySheet.getDataRange();
    var territory_max_rows = territory_range_data.getNumRows();

    for (var row_data=first_territroy_range_data_row; row_data<= territory_max_rows; row_data++) {
      row = territorySheet.getRange(row_data, publisher_name_column);
      if (! row.isBlank()) {
        row_value = row.getValue();
        
        if (re_publisher_spaces.test(row_value)) {
          Logger.log('Znalazłem spacje !! Dla: ' + row_value);
          new_value = row_value.replace(/\.\s+/, '.').trim();
          if (new_value != row_value){
            row.setValue(new_value);
            Logger.log('Po poprawie : ' + row_value);
          }
          fix_list.push('\nTeren nr : ' + number + ' '+ row_value);
        }
      };
    }
    
    if (number % 25 == 0 ) {
      source.toast(number);
    }
    // Testing condition in order to break loop
    if (number === 150) {
      break;
    }
  }
  ui.alert(fix_list);
}

function fixWrongNames() {
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput(""));
  var html_report = HtmlService.createTemplateFromFile('fixnames_sidebar');
  var first_territroy_range_data_row = 7;
  var publisher_name_column = 4;
  var html_result = [];
  var publishersSheet = source.getSheetByName('Głosiciele');
  var publishersSheet_rows = publishersSheet.getDataRange().getNumRows();
  var publishers_list_range = publishersSheet.getRange('D2:D' + publishersSheet_rows);
  var publishers_list = publishers_list_range.getValues();
  publishers_list.forEach(function (item, x, y) {
    publishers_list[x] = item[0];
  });
  
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    var territory_range_data = territorySheet.getDataRange();
    var territory_max_rows = territory_range_data.getNumRows();
    
    for (var row_data=first_territroy_range_data_row; row_data<= territory_max_rows; row_data++) {
      row = territorySheet.getRange(row_data, publisher_name_column);
      if (! row.isBlank()) {
        row_value = row.getValue();
        if (publishers_list.indexOf(row_value) < 0){
          html_result.push({'sheet_name': territorySheet.getName(),
                            'row': row_data,
                            'bad_entry': row_value});
        }
      };
    };
    
    if (number % 25 == 0 ) {
      source.toast(number);
    };
    // Testing condition in order to break loop
    if (number === 250) {
      break;
    };
  };
  html_report.results = html_result;
  ui.showSidebar(html_report.evaluate());
}

function open_tab_by_sheetname(sheet_name, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selected_sheet = ss.getSheetByName(sheet_name);
  selected_sheet.activate();
  if (row)
    selected_sheet.setActiveSelection('D'+row+':D'+row);
  return false;
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function territoryTables() {
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createTemplateFromFile("lista_terenow").evaluate());
}

function show_territory_cards_tools() {
  var ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createTemplateFromFile("s13_cardgen_sidebar").evaluate());
}

function generate_territory_cards() {
  var ui = SpreadsheetApp.getUi();
  var data = [];
  data.push("");
  source.toast("Start");
  for (var number = firstTerritory; number <= lastTerritory; number++) {
    var territorySheet = source.getSheetByName("Teren nr " + number);
    var territory_range_data = territorySheet.getDataRange();
    var territory_max_rows = territory_range_data.getNumRows();
    data.push(territorySheet.getRange("A7:D"+territory_max_rows).getValues());
    if (number % 25 == 0 ) {
      source.toast(number);
    };
  };
  source.toast("Ready");
  html_template = HtmlService.createTemplateFromFile("s13_cardgen");
  html_template.data = data;
  return html_template.evaluate().getContent();
}
