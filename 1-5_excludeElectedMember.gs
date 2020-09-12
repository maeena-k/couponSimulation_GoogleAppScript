function excludeElectedMember(spreadSheet, setSampleMemberSheet, index) {
  const ss = spreadSheet, set_sample_member_sheet = setSampleMemberSheet, loop_flag = index;
  
  ss.setActiveSheet(set_sample_member_sheet);
  const exclude_elected_member_sheet = ss.duplicateActiveSheet();
  exclude_elected_member_sheet.setName('1-5. exclude_reciepts_elected')
  exclude_elected_member_sheet.setFrozenRows(1);
  
  if (loop_flag >= 3) {
    let last_row = exclude_elected_member_sheet.getLastRow();
    
    const formula_value = "=countifs('COUPON_TARGET_USER_LIST'!A:A,A2,'COUPON_TARGET_USER_LIST'!B:B,B2)";
    exclude_elected_member_sheet.getRange('C2').setFormula(formula_value);
    exclude_elected_member_sheet.getRange('C3:C'+last_row).setFormula(exclude_elected_member_sheet.getRange('C2').getFormula());
    
    const copy_range = exclude_elected_member_sheet.getRange('C2:C'+last_row);
    copy_range.setValues(copy_range.getValues());
    
    exclude_elected_member_sheet.getRange('A:C').sort([{column:3, ascending:false}]);
    const flag_list = copy_range.getValues();
    const flag_list_len = flag_list.length;
    
    if (flag_list_len >= 1) {
      for (let i = 0; i <= flag_list_len;i++){
        if (flag_list[i] === 0){
          exclude_elected_member_sheet.deleteRows(2, i);
          break;
        }
      }
    }
    last_row = exclude_elected_member_sheet.getLastRow();
    exclude_elected_member_sheet.getRange('C2:C'+last_row).clearContent();
    exclude_elected_member_sheet.getRange('A:C').sort([{column:1, ascending:true}, {column:2, ascending:true}]);
  }
  
  return exclude_elected_member_sheet;
}