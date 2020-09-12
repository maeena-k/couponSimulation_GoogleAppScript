function deleteSampleSheetsUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coupon_target_member_sheet = ss.getSheetByName('COUPON_TARGET_USER_LIST');
  
  let keep_list = ['campaign_outline', 'election_target_builder', 'coupon_list_builder',
                 'EXECUTE_SIMULATION_BUTTON', 'COUPON_TARGET_USER_LIST', '==> Simulation Sample Sheets'];
  
  deleteSampleSheets(ss, keep_list);
  clearCouponTargetMember(coupon_target_member_sheet);
}