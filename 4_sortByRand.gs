function sortByRand(spreadSheet, calMinRandMemberSheet) {
  const ss = spreadSheet, cal_min_rand_member_sheet = calMinRandMemberSheet;
  
  ss.setActiveSheet(cal_min_rand_member_sheet);
  const sort_rand_member_sheet = ss.duplicateActiveSheet();
  
  sort_rand_member_sheet.setName('4. sort_by_rand_asc');
  sort_rand_member_sheet.setFrozenRows(1);
  
  sort_rand_member_sheet.getRange('A:C').sort({column:3,ascending:true});
  
  return sort_rand_member_sheet;
}