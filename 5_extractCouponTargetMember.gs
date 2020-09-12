function extractCouponTargetMember(spreadSheet, sortRandMemberSheet, targetCouponNum) {
  const ss = spreadSheet, sort_rand_member_sheet = sortRandMemberSheet, target_coupon_num = targetCouponNum;
  
  ss.setActiveSheet(sort_rand_member_sheet);
  const extract_coupon_member_sheet = ss.duplicateActiveSheet();
  
  extract_coupon_member_sheet.setName('5. extract_target_user');
  extract_coupon_member_sheet.setFrozenRows(1);
  
  const last_row = extract_coupon_member_sheet.getLastRow();
  extract_coupon_member_sheet.getRange('A'+(target_coupon_num+2)+':C'+last_row).clearContent();
  
  return extract_coupon_member_sheet;
}