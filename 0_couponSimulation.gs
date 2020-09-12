function couponSimulation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const set_target_member_sheet = ss.getSheetByName('election_target_builder');
  const set_target_coupon_sheet = ss.getSheetByName('coupon_list_builder');
  const coupon_target_member_sheet = ss.getSheetByName('COUPON_TARGET_USER_LIST');
  const coupon_count = set_target_coupon_sheet.getLastRow();
  
  let keep_list = ['campaign_outline', 'election_target_builder', 'coupon_list_builder',
                   'EXECUTE_SIMULATION_BUTTON', 'COUPON_TARGET_USER_LIST', '==> Simulation Sample Sheets'];
  let set_sample_member_sheet, exclude_elected_member_sheet, set_rand_member_sheet,
      cal_min_rand_member_sheet, sort_rand_member_sheet, extract_coupon_member_sheet;
  
  for (let i = 2; i <= coupon_count; i++) {
    const target_coupon_num = set_target_coupon_sheet.getRange(i, 4).getValue();
    const target_coupon_name = set_target_coupon_sheet.getRange(i, 2).getValue();
    
    /* PRE-STEP */
    deleteSampleSheets(ss, keep_list);
    
    if (i == 2){
      clearCouponTargetMember(coupon_target_member_sheet);
    }
    
    /* STEP1 */
    if (i == 2) {
      set_sample_member_sheet = createTargetMember(ss, set_target_member_sheet);
      keep_list.push('1. election_target_users');
    } else {
      set_sample_member_sheet = ss.getSheetByName('1. election_target_users');
    }
    
    /* STEP1.5 */
    exclude_elected_member_sheet = excludeElectedMember(ss, set_sample_member_sheet, i);
    
    /* STEP2 */
    set_rand_member_sheet = setRandNum(ss, exclude_elected_member_sheet);
    
    /* STEP3 */
    cal_min_rand_member_sheet = calMinRand(ss, set_rand_member_sheet);
    
    /* STEP4 */
    sort_rand_member_sheet = sortByRand(ss, cal_min_rand_member_sheet);
    
    /* STEP5 */
    extract_coupon_member_sheet = extractCouponTargetMember(ss, sort_rand_member_sheet, target_coupon_num);
    
    /* STEP6 */
    saveCouponTargetMember(ss, extract_coupon_member_sheet, coupon_target_member_sheet, target_coupon_name);
  }
}