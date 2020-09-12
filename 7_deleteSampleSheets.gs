function deleteSampleSheets(spreadSheet, keepList) {
  const ss = spreadSheet, keep_list = keepList;
  
  const sheet_count = ss.getNumSheets();
  let target_sheet;
  
  for (let i = sheet_count-1; i >= 0; i--){
    target_sheet = ss.getSheets()[i];
    
    if (typeof(target_sheet.getName()) == 'undefined' || keep_list.indexOf(target_sheet.getName()) < 0){
      ss.deleteSheet(target_sheet);
    }
  }
}