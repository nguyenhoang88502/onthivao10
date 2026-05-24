// ==========================================
// MASTER SEED: All 30 tests in 1 function
// Run: seedAll30Tests()
// Then call each batch: seed01_05(), seed06_10(), etc.
// ==========================================

var SS_ID = 'YOUR_SPREADSHEET_ID_HERE';

function getSS_() { return SpreadsheetApp.openById(SS_ID); }

function addQ_(tid, qs) {
  var sheet = getSS_().getSheetByName('Questions');
  qs.forEach(function(q, i) {
    var p = (i < 4) ? 1 : (i < 16) ? 2 : (i < 28) ? 3 : 4;
    sheet.appendRow(['T'+tid+'_Q'+(i+1), tid, p, q[0], q[1], q[2]||'', q[3]||'', q[4]||'', q[5]||'', q[6]||'', q[7], q[8]||0.25, q[9]||'']);
  });
}

function seedAll30Tests() {
  var ss = getSS_();
  var ts = ss.getSheetByName('Tests');
  var qs = ss.getSheetByName('Questions');
  ts.clearContents(); qs.clearContents();
  ts.appendRow(['test_id','title','description','time_limit']);
  qs.appendRow(['question_id','test_id','part_number','section_title','question_text','context_image_url','option_a','option_b','option_c','option_d','correct_answer','points','explanation_template']);

  var titles = [
    'De thi thu so 01|Bam sat de minh hoa So GD&DT TPHCM 2025',
    'De thi thu so 02|Chu de Moi truong & Nang luong xanh',
    'De thi thu so 03|Chu de Giao duc & Cong nghe 4.0',
    'De thi thu so 04|Chu de Gia dinh & Ky nang song',
    'De thi thu so 05|Chu de Van hoa & Le hoi Viet Nam',
    'De thi thu so 06|Chu de Suc khoe & Dinh duong',
    'De thi thu so 07|Chu de Giao thong & An toan',
    'De thi thu so 08|Trong tam thi & ngu phap nang cao',
    'De thi thu so 09|Chu de Khoa hoc & Kham pha',
    'De thi thu so 10|Chu de Am nhac & Giai tri',
    'De thi thu so 11|Chu de The thao & Thanh tich',
    'De thi thu so 12|Chu de Du lich & Kham pha',
    'De thi thu so 13|Chu de Thoi trang & Phong cach',
    'De thi thu so 14|Chu de Nghe nghiep & Tuong lai',
    'De thi thu so 15|Chu de Tinh nguyen & Cong dong',
    'De thi thu so 16|Bam sat de chinh thuc 2024',
    'De thi thu so 17|Chu de Bien doi khi hau',
    'De thi thu so 18|Chu de An toan thuc pham',
    'De thi thu so 19|Chu de Tri tue nhan tao (AI)',
    'De thi thu so 20|Chu de Mang xa hoi & Truyen thong',
    'De thi thu so 21|Chu de Lich su & Di tich',
    'De thi thu so 22|Chu de Thien nhien & Dong vat',
    'De thi thu so 23|Chu de Sach & Van hoa doc',
    'De thi thu so 24|Chu de Nang luong tai tao',
    'De thi thu so 25|Chu de Do thi thong minh',
    'De thi thu so 26|Bam sat de chinh thuc 2023',
    'De thi thu so 27|Chu de Giao duc STEM',
    'De thi thu so 28|Chu de Ky nang mem & Lanh dao',
    'De thi thu so 29|Chu de Bao ve dong vat quy hiem',
    'De thi thu so 30|Tong on tap cuoi cung'
  ];
  titles.forEach(function(t, i) {
    var parts = t.split('|');
    ts.appendRow([i+1, parts[0], parts[1], 90]);
  });

  seed01_05(); seed06_10(); seed11_15();
  seed16_20(); seed21_25(); seed26_30();
  Logger.log('All 30 tests = 1200 questions seeded!');
}

// ===== BATCH FILES =====
// Each batch function is in its own .gs file for size management.
// Copy all batch files into GAS Editor and run seedAll30Tests().
