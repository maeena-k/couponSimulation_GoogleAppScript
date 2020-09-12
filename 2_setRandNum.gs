class Random {
  constructor(seed = 88675123) {
    this.x = 123456789;
    this.y = 362436069;
    this.z = 521288629;
    this.w = seed;
  }
  
  next() {
    let t;
 
    t = this.x ^ (this.x << 11);
    this.x = this.y; this.y = this.z; this.z = this.w;
    return this.w = (this.w ^ (this.w >>> 19)) ^ (t ^ (t >>> 8));
  }
  
  nextInt(min, max) {
    const r = Math.abs(this.next());
    return min + (r % (max + 1 - min));
  }
}

function setRandNum(spreadSheet, excludeElectedMemberSheet, index) {
  const ss = spreadSheet, exclude_elected_member_sheet = excludeElectedMemberSheet;
  const seed = index;
  
  ss.setActiveSheet(excludeElectedMemberSheet);
  const set_rand_member_sheet = ss.duplicateActiveSheet();
  
  set_rand_member_sheet.setName('2. set_rand_num');
  set_rand_member_sheet.getRange('C1').setValue('rand_num');
  
  const reciept_num = set_rand_member_sheet.getLastRow() - 1;
  const random = new Random(seed);
  
  let num, i = 2, num_list = [], temp_num_list = [];
  while (i <= reciept_num+1) {
    num = random.nextInt(1,reciept_num);
    if (temp_num_list.indexOf(num) < 0) {
      temp_num_list.push(num);
      num_list.push([num]);
      i++;
    }
  }
  set_rand_member_sheet.getRange('C2:C'+(reciept_num+1)).setValues(num_list);
  return set_rand_member_sheet;
}