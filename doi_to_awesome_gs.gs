function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Options");
  var items = menu.addItem("Get PDF ID", "getPdfId");
  var item2 = menu.addItem("Set owner to MVandenberg", "changeOwner");
  var item3 = menu.addItem("Add to Publications by IRRI Staff folder", "addToIRRIPubs");
  var item4 = menu.addItem("Add to Open Access folder", "addToOpenAccess");
  items.addToUi();
  item2.addToUi();
  item3.addToUi();
  item4.addToUi();
}

function changeOwner() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pdfId = ss.getRange("D3").getValue();
  var thePdf = DriveApp.getFileById(pdfId);
  var mVandenberg = thePdf.setOwner("m.vandenberg@irri.org");
  //var mBonador = thePdf.setOwner("m.bonador@irri.org");
}

function addToIRRIPubs() {
  var irriPubFolder = DriveApp.getFolderById("0B-Hz5UNUXAiWdmF3SFdqRl92UFE");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pdfId = ss.getRange("D3").getValue();
  var thePdf = DriveApp.getFileById(pdfId);
  irriPubFolder.addFile(thePdf)
}

function addToOpenAccess() {
  var openAccess = DriveApp.getFolderById("0B-Hz5UNUXAiWV0E5XzZfYXdGTFU");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pdfId = ss.getRange("D3").getValue();
  var thePdf = DriveApp.getFileById(pdfId);
  openAccess.addFile(thePdf)
}

function getPdfId(){
  //var folders = DriveApp.getFolders(); // will search all folders of your shared google drive but will take forever, update the years below or write a script that will 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year_value = ss.getRange("A3").getValue();
  var folders = DriveApp.getFoldersByName(year_value);
  
  while(folders.hasNext()){
    var folder = folders.next();
    var files = folder.getFiles();
    while(files.hasNext()){
      var file = files.next();
      var the_file = file.getName();      
      var title_value = ss.getRange("F3").getValue();
      if(file == the_file) {
        if(the_file.indexOf(title_value)>= 0){
          var pdf_Id = file.getId();
          Logger.log(pdf_Id);
        }
      }
      var sheet = ss.getSheets()[0];
      var cell_d = sheet.getRange("D3");
      cell_d.setValue(pdf_Id);
      }    
  }
}

function onEdit(evt){
  Logger.log('Something was edited, previous value was: ' + evt.oldValue);
  var column = evt.range.getColumn();
  var row = evt.range.getRow();
  var cell = evt.range.getA1Notation();
  Logger.log(column + " and " + row);
  Logger.log(cell);
  var doi = SpreadsheetApp.getActiveSheet().getRange(cell).getValue()
  Logger.log(doi);
  getAltmetricId(doi);
  doi_fetch(doi);  
  SpreadsheetApp.getActiveSheet().getRange(row, column).clearContent()
}

function getAltmetricId(the_doi) {
  var options = {
    'muteHttpExceptions' : true
  };
  var AMUrl = "http://api.altmetric.com/v1/doi/" + the_doi;
  var response = UrlFetchApp.fetch(AMUrl,options);
  var content = response.getContentText();  
  if(content == "Not Found"){
    return false
  } else {
    var content = response.getContentText();
  }
   
  var json = JSON.parse(content);
  Logger.log(json)
  var altmetric_id = json['altmetric_id'];
  var formula = '=if(O3="","",iferror(getAltmetricsScore(O3),""))';  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cell_o = sheet.getRange('O3');
  var cell_p = sheet.getRange('P3');    
  cell_o.setValue(altmetric_id);
  cell_p.setValue(formula) ;
  Logger.log(altmetric_id);
  
}

function doi_fetch(the_doi) {
  //var the_doi = '10.3390/cli7040048' //published-online
  //var the_doi = '10.1016/j.agwat.2019.02.037' //published-print
  //var the_doi = '10.1128/JVI.01503-18' //no issue element
  //var the_doi = '10.3390/genes10040290'
  //var the_doi = '10.1093/bioinformatics/btz190';
  //var the_doi = '10.1093/gbe/evz084'; - no source
  var url_string = "https://api.crossref.org/v1/works/" + the_doi;
  var response = UrlFetchApp.fetch(url_string);
  var content = response.getContentText();
  var json = JSON.parse(content);
    
  var jsonKeys = Object.keys(json);
  var year = '';
  var published_print_date = json['message']['published-print'];
  var published_online_date = json['message']['published-online'];
  for(var y in json['message']){
    if(json['message'].hasOwnProperty('published-print')){
      year = published_print_date['date-parts'][0][0];
    } else {
      year = published_online_date['date-parts'][0][0];
    }
  } 
   
  
  var authors = json['message']['author'];
  var all_authors = [];
  for(var x in authors){
    if(!('given' in authors[x])){
      all_authors.push(authors[x]['family']);
    } else if(authors[x]['sequence'] == 'first'){
      all_authors.push(authors[x]['family'] + ', ' + authors[x]['given']);
    } else if(authors[x]['sequence'] == 'additional'){
      all_authors.push(authors[x]['given'] + ' ' + authors[x]['family']);
    } 
  } 
    
  if(all_authors.length > 1){
    all_authors.splice(all_authors.length-1, 0, 'and');
    var last_two_items = all_authors.slice(all_authors.length-2);
    last_two_items = last_two_items.join(' ');
    all_authors.splice(all_authors.length-2, all_authors.length, last_two_items);
  }
  all_authors = all_authors.join(', ');
    
  
  var title = json['message']['title'];
  var book_source = json['message']['container-title'][0];
  var journal = json['message']['short-container-title'][0];
  var volume = json['message']['volume'];
  var issue = json['message']['issue'];
  var type = json['message']['type'];
  var page = json['message']['page'];
  var source = [];
  var link_formula = '=if(D3="","","<a href=" & char(34) & if(H3<>"",IF(J3<>"","http://dx.doi.org/ "& J3,"https://docs.google.com/a/irri.org/file/d/" & D3 & "/view"),"https://docs.google.com/a/irri.org/file/d/" & D3 & "/view") & char(34) & " target=" & char(34) & "_blank" & char(34) & "><img src=" & char(34) & if(H3="","https://sites.google.com/a/irri.org/publications-by-irri-staff/_/rsrc/1429577224651/2014/requestaccess.png","https://sites.google.com/a/irri.org/publications-by-irri-staff/2014/openaccess.png") & char(34) & "></a>")';
  var impact_factor = '=if(L3="","",iferror(if(upper(vlookup(L3,\'2016ImpactFactors\'!A:B,1,False))=upper(L3),vlookup(L3,\'2016ImpactFactors\'!A:B,2,False),"No"),"Not found"))';
  var doi_display = '=if(J3="",""," <a href=" & char(34) & "http://dx.doi.org/" & J3 & char(34) & " target=" & char(34) & "_blank" & char(34) &">" & "doi:" & J3 & "</a>"  & if(O3<>""," <a href=" & char(34) & "https://www.altmetric.com/details/" & O3 & char(34) & " target=" & char(34) & "_blank" & char(34) &">" & "<button type=button><b>Altmetrics" & if(P3="",""," score " & P3) & "</b></button>" & "</a>",""))';
  
  if(type == "book-chapter"){
    source = "In: " + book_source.concat(' p. ', page);
    journal = '';
    type = 'Book'
  } else if(type == "journal-article"){
    if (json['message'].hasOwnProperty('volume')==false && json['message'].hasOwnProperty('issue')==false && json['message'].hasOwnProperty('page')==false) {
      source = journal
      type = 'Journal'
    } else if(json['message'].hasOwnProperty('issue')==false && json['message'].hasOwnProperty('page')==false){
      source = journal.concat(' ', volume);
      type = 'Journal'
    } else if(json['message'].hasOwnProperty('volume')==false){
      source = journal.concat(' p. ', page);
      type = 'Journal'
    } else if (json['message'].hasOwnProperty('issue')==false){
      source = journal.concat(' ', volume, ' p. ', page);
      type = 'Journal'
    } else if(json['message'].hasOwnProperty('page')==false){
      source = journal.concat(' ', volume,', no. ', issue);
      type = 'Journal'
    } else {
      source = journal.concat(', ', volume,' no. ', issue, ' p. ', page);
      type = 'Journal'
    }
  } 
  
  var d_o_i = json['message']['DOI'];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var cell_a = sheet.getRange('A3');
  var cell_b = sheet.getRange('B3');
  var cell_e = sheet.getRange('E3');
  var cell_f = sheet.getRange('F3');
  var cell_g = sheet.getRange('G3');
  var cell_h = sheet.getRange('H3');
  var cell_i = sheet.getRange('I3');
  var cell_j = sheet.getRange('J3');
  var cell_k = sheet.getRange('K3');
  var cell_l = sheet.getRange('L3');
  var cell_m = sheet.getRange('M3');
  
  cell_a.setValue(year);
  cell_b.setValue(link_formula);
  cell_e.setValue(all_authors);
  cell_f.setValue(title);
  cell_g.setValue(source);
  cell_i.setValue(type);
  cell_j.setValue(d_o_i);
  cell_k.setValue(impact_factor);
  cell_l.setValue(journal);
  cell_m.setValue(doi_display);
  
  Logger.log(title)
  
  }

function getAltmetricsScore(AMID) {
  var options = {
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch('https://www.altmetric.com/details/' + AMID,options);
  htmltxt = response.getContentText();
  return +htmltxt.substring(htmltxt.indexOf("score=")+6,htmltxt.indexOf("&types="));  
}