function calMinRand(spreadSheet, setRandMemberSheet) {
  const ss = spreadSheet, set_rand_member_sheet = setRandMemberSheet;
  
  const sheet_count = ss.getNumSheets();
  const cal_min_rand_member_sheet = ss.insertSheet(sheet_count);
  
  cal_min_rand_member_sheet.setName('3. calculate_min_rand_num');
  cal_min_rand_member_sheet.getRange('A1').setValue('member_id');
  cal_min_rand_member_sheet.getRange('B1').setValue('reciept_num');
  cal_min_rand_member_sheet.getRange('C1').setValue('rand_num_min');
  cal_min_rand_member_sheet.setFrozenRows(1);
  
  const last_row_rand = set_rand_member_sheet.getLastRow();
  let member_list = set_rand_member_sheet.getRange('A2:A'+last_row_rand).getValues();
  
  let uniq_member_list = [], temp_member_list = [];
  let member_id;
  
  for (let i = 0; i < member_list.length; i++) {
    member_id = member_list[i][0];
    
    if (temp_member_list.indexOf(member_id) < 0) {
      temp_member_list.push(member_id);
      uniq_member_list.push([member_id]);
    }
  }
  const target_member_num = uniq_member_list.length;
  const formula_rand = "=Minifs('2. set_rand_num'!C:C, '2. set_rand_num'!A:A,A2)",
      formula_reciept = "=index('2. set_rand_num'!A:C,match(C2, '2. set_rand_num'!C:C, 0), 2)";
  
  cal_min_rand_member_sheet.getRange(2, 1, target_member_num, 1).setValues(uniq_member_list);
  cal_min_rand_member_sheet.getRange('C2').setFormula(formula_rand);
  cal_min_rand_member_sheet.getRange('B2').setFormula(formula_reciept);
  
  const last_row_rand_min = cal_min_rand_member_sheet.getLastRow();
  
  cal_min_rand_member_sheet.getRange('B2:B'+last_row_rand_min).setFormula(cal_min_rand_member_sheet.getRange('B2').getFormula());
  cal_min_rand_member_sheet.getRange('C2:C'+last_row_rand_min).setFormula(cal_min_rand_member_sheet.getRange('C2').getFormula());
  
  const copy_range = cal_min_rand_member_sheet.getRange('B2:C'+last_row_rand_min);
  copy_range.setValues(copy_range.getValues());
  
  return cal_min_rand_member_sheet;
}