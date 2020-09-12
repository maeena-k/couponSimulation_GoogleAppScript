function clearCouponTargetMember(couponTargetMemberSheet) {
  const coupon_target_member_sheet = couponTargetMemberSheet;
  const last_row = coupon_target_member_sheet.getLastRow();
  
  if (last_row >= 2) {
    coupon_target_member_sheet.getRange('A2:C' + last_row).clearContent();
  }
}