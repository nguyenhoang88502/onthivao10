# Huong dan cai dat - He thong Luyen thi vao 10 TPHCM Mon Tieng Anh

## 1. Cau truc du an

```
Code.gs     -> Google Apps Script backend (API + xu ly du lieu)
Index.html  -> Frontend giao dien nguoi dung (Tailwind CSS + Vanilla JS)
SETUP.md    -> Tai lieu huong dan nay
```

## 2. Tao Google Sheets co so du lieu

### Buoc 1: Tao Spreadsheet moi
Truy cap https://sheets.google.com va tao mot spreadsheet moi.

### Buoc 2: Tao 3 sheet (tab) voi cac cot sau

#### Sheet `Tests`
| test_id | title | description | time_limit |
|---------|-------|-------------|------------|
| 1 | De thi thu so 1 | Bam sat cau truc de minh hoa So GD&DT TPHCM 2025 | 90 |
| 2 | De thi thu so 2 | Trong tam ngu phap hoc ky 2 - Co cau hoi bien bao | 90 |
| ... | ... | ... | 90 |

**test_id**: So nguyen, duy nhat (1-30)
**title**: Ten de thi
**description**: Mo ta ngan ve de thi
**time_limit**: Thoi gian lam bai (phut), thong thuong la 90

#### Sheet `Questions`
| question_id | test_id | part_number | section_title | question_text | context_image_url | option_a | option_b | option_c | option_d | correct_answer | points | explanation_template |
|-------------|---------|-------------|---------------|---------------|-------------------|----------|----------|----------|----------|----------------|--------|---------------------|

**question_id**: Chuoi duy nhat, VD: "T1_Q1" (Test 1, Question 1)
**test_id**: So nguyen, tham chieu den Tests.test_id
**part_number**: 1-4 (1=Phonetics, 2=Vocab/Grammar/Comm, 3=Reading, 4=Writing)
**section_title**: Ten phan nho, VD: "Pronunciation: -ed endings", "Cloze Test", "Word Form"
**question_text**: Noi dung cau hoi (tieng Anh)
**context_image_url**: URL hinh anh bien bao / bieu do (de trong neu khong co)
**option_a..d**: Cac dap an lua chon (de trong neu la cau tu luan Part 4)
**correct_answer**: Dap an dung (A/B/C/D hoac cau tra loi day du cho Part 4)
**points**: Diem so cho cau nay (thong thuong 0.25)
**explanation_template**: Giai thich mau (co the de trong, AI se tu sinh)

**QUAN TRONG VE PHAN BO CAU HOI (40 cau / de):**
- Part 1 (Phonetics): **4 cau** (Q1-Q4)
- Part 2 (Vocab, Grammar & Communication): **12 cau** (Q5-Q16)
- Part 3 (Reading Comprehension): **12 cau** (Q17-Q28) - gom 6 Cloze Test + 6 Reading
- Part 4 (Writing): **12 cau** (Q29-Q40) - gom 6 Word Form + 2 Dictionary + 4 Transformation

#### Sheet `Submissions`
Sheet nay tu dong duoc ghi du lieu khi hoc sinh nop bai. Khong can nhap du lieu thu cong.

| submission_id | student_name | student_id | test_id | answers_json | score | timestamp |
|---------------|--------------|------------|---------|--------------|-------|-----------|

## 3. Cai dat Google Apps Script

### Buoc 1: Mo Google Apps Script Editor
Tu Google Sheets: **Extensions > Apps Script**

### Buoc 2: Copy code
1. Xoa code mac dinh trong editor
2. Copy toan bo noi dung tu file `Code.gs` va dan vao
3. Tao file HTML moi: **File > New > HTML file**, dat ten la `Index`
4. Copy toan bo noi dung tu file `Index.html` va dan vao file `Index.html` vua tao

### Buoc 3: Cap nhat SPREADSHEET_ID
Tim dong sau trong `Code.gs` va thay YOUR_SPREADSHEET_ID_HERE bang ID cua Google Sheets:

```javascript
SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
```

**Cach lay Spreadsheet ID**: URL cua Google Sheets co dang:
`https://docs.google.com/spreadsheets/d/XXXXX/edit`
`XXXXX` chinh la SPREADSHEET_ID.

### Buoc 4: Cai dat DeepSeek API Key
1. Dang ky tai khoan va lay API key tai https://platform.deepseek.com
2. Trong GAS Editor: **Project Settings > Script Properties**
3. Them property moi:
   - Property: `DEEPSEEK_API_KEY`
   - Value: `sk-your-deepseek-api-key-here`

### Buoc 5: Deploy Web App
1. Nhan **Deploy > New deployment**
2. Chon type: **Web app**
3. Description: `He thong luyen thi vao 10 TPHCM`
4. Execute as: **Me** (your account)
5. Who has access: **Anyone** (hoac Anyone within your organization)
6. Nhan **Deploy**
7. Cap quyen truy cap (Authorize access)
8. Copy URL duoc cung cap - do la URL cua web app

### Buoc 6: Kiem tra
Mo URL web app trong trinh duyet de kiem tra:
- Dashboard hien thi danh sach de thi
- Nhan "Bat dau" de vao giao dien lam bai
- Giao dien lam bai co dong ho dem nguoc 90:00
- Nop bai de xem ket qua
- Nhan "Giai thich bang AI" de xem DeepSeek giai thich

## 4. Nhap du lieu de thi mau

### Cach 1: Nhap truc tiep vao Google Sheets
Nhap tung dong du lieu vao sheet `Tests` va `Questions` theo cau truc ben tren.

### Cach 2: Su dung Google Apps Script de sinh du lieu mau
Chay ham sau trong GAS Editor de tao du lieu mau cho Test 1:

```javascript
function seedSampleTest1() {
  var ss = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID_HERE');
  var testsSheet = ss.getSheetByName('Tests');
  var questionsSheet = ss.getSheetByName('Questions');

  // Them test
  testsSheet.appendRow([1, 'De thi thu so 1', 'Bam sat cau truc de minh hoa So GD&DT TPHCM 2025', 90]);

  var questions = [
    // Part 1: Phonetics (4 cau)
    [1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. watched B. stopped C. needed D. walked', '', 'A. watched', 'B. stopped', 'C. needed', 'D. walked', 'C', 0.25, ''],
    [1, 'Pronunciation: Vowels', 'Choose the word whose underlined part is pronounced differently: A. seat B. great C. meat D. beat', '', 'A. seat', 'B. great', 'C. meat', 'D. beat', 'B', 0.25, ''],
    [1, 'Word Stress: 2-syllable', 'Choose the word with different stress pattern: A. mother B. father C. believe D. brother', '', 'A. mother', 'B. father', 'C. believe', 'D. brother', 'C', 0.25, ''],
    [1, 'Word Stress: 3-syllable', 'Choose the word with different stress pattern: A. family B. computer C. holiday D. cinema', '', 'A. family', 'B. computer', 'C. holiday', 'D. cinema', 'B', 0.25, ''],
    // Part 2: Vocabulary, Grammar & Communication (12 cau)
    [2, 'Vocabulary in Context', 'The government has introduced a new ____ to protect the environment. A. law B. rule C. policy D. regulation', '', 'A. law', 'B. rule', 'C. policy', 'D. regulation', 'C', 0.25, ''],
    [2, 'Grammar: Tenses', 'She ____ English for five years before she moved to London. A. has studied B. had studied C. was studying D. studied', '', 'A. has studied', 'B. had studied', 'C. was studying', 'D. studied', 'B', 0.25, ''],
    // ... (them 10 cau nua cho Part 2)
    // Part 3: Reading Comprehension (12 cau)
    // Part 4: Writing (12 cau)
  ];

  questions.forEach(function(q, i) {
    var partNum = (i < 4) ? 1 : (i < 16) ? 2 : (i < 28) ? 3 : 4;
    questionsSheet.appendRow(['T1_Q' + (i+1), 1, partNum, q[0], q[1], q[2], q[3], q[4], q[5], q[6], q[7], q[8], q[9]]);
  });
}
```

## 5. Luu y quan trong

### Bao mat
- **Dap an dung (`correct_answer`) khong bao gio duoc gui xuong client** truoc khi nop bai. Ham `getTestDetails()` loai bo truong nay.
- **DeepSeek API key chi ton tai o server** trong Script Properties, khong lo ra client.
- Su dung `LockService` de tranh xung dot khi nhieu hoc sinh nop bai cung luc.

### Gioi han cua GAS
- Google Apps Script co gioi han 6 phut thuc thi. Cac ham hien tai deu chay duoi gioi han nay.
- UrlFetchApp co gioi han 20,000 requests/ngay. Dieu nay du cho viec goi DeepSeek API.
- HtmlService tu dong san loc HTML de bao ve. Da them `.setXFrameOptionsMode(ALLOWALL)`.

### Tuy chinh
- Thay doi `TIME_LIMIT_MINUTES` trong `Code.gs` de dieu chinh thoi gian lam bai.
- Thay doi `POINTS_PER_QUESTION` de dieu chinh cach tinh diem.
- Thay doi `DEEPSEEK_MODEL` de su dung model DeepSeek khac (VD: deepseek-reasoner).
- Dieu chinh `systemPrompt` trong `callDeepSeekExplanation()` de thay doi giong dieu AI.

### Kiem tra loi
Tat ca cac ham deu co try-catch va tra ve thong bao loi ro rang bang tieng Viet:
- "Không thể tải danh sách đề thi. Vui lòng thử lại."
- "Hệ thống AI đang quá tải. Vui lòng thử lại sau vài giây."
- "Không thể nộp bài. Vui lòng thử lại."

## 6. Thong tin lien he
He thong duoc xay dung cho ky thi Tuyen sinh vao lop 10 tai TPHCM.
Dinh dang de thi tuan thu cau truc chinh thuc cua So GD&DT TPHCM.
