function createTargetMember(spreadSheet, setTargetMemberSheet) {
  const ss = spreadSheet, set_target_member_sheet = setTargetMemberSheet;
  
  const sheet_count = ss.getNumSheets();
  const set_sample_member_sheet = ss.insertSheet(sheet_count);
  
  set_sample_member_sheet.setName('1. election_target_users')
  set_sample_member_sheet.getRange('A1').setValue('member_id')
  set_sample_member_sheet.getRange('B1').setValue('reciept_num')
  set_sample_member_sheet.setFrozenRows(1);
  
  const last_row = set_target_member_sheet.getLastRow();
  var reciept_num, member_num;
  
  let j = 1, paste_row_num = 2;
  
  for (let i = 3; i <= last_row; i++) {
    reciept_num = set_target_member_sheet.getRange('A' + i).getValue();
    member_num = set_target_member_sheet.getRange('B' + i).getValue();
    
    for (let k = 1; k <= member_num; k++){
      for (let m = 1; m <= reciept_num; m++){
        set_sample_member_sheet.getRange('A' + paste_row_num).setValue(j);
        set_sample_member_sheet.getRange('B' + paste_row_num).setValue(m);
        paste_row_num++;
      }
      j++;
    }
  }
  return set_sample_member_sheet;
}