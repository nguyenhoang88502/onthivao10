// ==========================================
// MASTER SEED: All 25 Tests - Run seedAll25Tests()
// This is the ONLY function you need to run.
// It delegates to the batch files for tests 1-25.
// ==========================================

function seedAll25Tests() {
  var ss = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID_HERE');
  var testsSheet = ss.getSheetByName('Tests');
  var questionsSheet = ss.getSheetByName('Questions');

  testsSheet.clearContents();
  questionsSheet.clearContents();
  testsSheet.appendRow(['test_id','title','description','time_limit']);
  questionsSheet.appendRow(['question_id','test_id','part_number','section_title','question_text','context_image_url','option_a','option_b','option_c','option_d','correct_answer','points','explanation_template']);

  var tests = [
    { id:1,  title:'Đề thi thử số 01', desc:'Bám sát đề minh họa Sở GD&ĐT TPHCM năm 2025 - Cấu trúc chuẩn 4 phần', time:90 },
    { id:2,  title:'Đề thi thử số 02', desc:'Chủ đề Môi trường & Năng lượng xanh - Đọc hiểu về biến đổi khí hậu', time:90 },
    { id:3,  title:'Đề thi thử số 03', desc:'Chủ đề Giáo dục & Công nghệ 4.0 - Từ vựng học thuật và giao tiếp', time:90 },
    { id:4,  title:'Đề thi thử số 04', desc:'Chủ đề Gia đình & Kỹ năng sống - Ngữ pháp trọng tâm câu điều kiện', time:90 },
    { id:5,  title:'Đề thi thử số 05', desc:'Chủ đề Văn hóa & Lễ hội Việt Nam - Đọc hiểu về Tết Nguyên Đán', time:90 },
    { id:6,  title:'Đề thi thử số 06', desc:'Chủ đề Sức khỏe & Dinh dưỡng - Từ vựng y tế và lối sống lành mạnh', time:90 },
    { id:7,  title:'Đề thi thử số 07', desc:'Chủ đề Giao thông & An toàn - Biển báo và tình huống thực tế', time:90 },
    { id:8,  title:'Đề thi thử số 08', desc:'Trọng tâm thì & ngữ pháp nâng cao - Câu bị động và câu tường thuật', time:90 },
    { id:9,  title:'Đề thi thử số 09', desc:'Chủ đề Khoa học & Khám phá - Đọc hiểu về phát minh và nhà khoa học', time:90 },
    { id:10, title:'Đề thi thử số 10', desc:'Chủ đề Âm nhạc & Giải trí - Từ vựng về nghệ thuật và biểu diễn', time:90 },
    { id:11, title:'Đề thi thử số 11', desc:'Chủ đề Thể thao & Thành tích - Đọc hiểu về Olympic và thể thao', time:90 },
    { id:12, title:'Đề thi thử số 12', desc:'Chủ đề Du lịch & Khám phá - Từ vựng về địa danh và hành trình', time:90 },
    { id:13, title:'Đề thi thử số 13', desc:'Chủ đề Thời trang & Phong cách - Mệnh đề quan hệ và câu so sánh', time:90 },
    { id:14, title:'Đề thi thử số 14', desc:'Chủ đề Nghề nghiệp & Tương lai - Ước mơ và định hướng nghề nghiệp', time:90 },
    { id:15, title:'Đề thi thử số 15', desc:'Chủ đề Tình nguyện & Cộng đồng - Đọc hiểu về hoạt động xã hội', time:90 },
    { id:16, title:'Đề thi thử số 16', desc:'Bám sát đề chính thức 2024 - Đà Nẵng và tiết kiệm năng lượng', time:90 },
    { id:17, title:'Đề thi thử số 17', desc:'Chủ đề Biến đổi khí hậu - Từ vựng môi trường và giải pháp xanh', time:90 },
    { id:18, title:'Đề thi thử số 18', desc:'Chủ đề An toàn thực phẩm - Đọc hiểu về nông nghiệp hữu cơ', time:90 },
    { id:19, title:'Đề thi thử số 19', desc:'Chủ đề Trí tuệ nhân tạo (AI) - Từ vựng công nghệ tương lai', time:90 },
    { id:20, title:'Đề thi thử số 20', desc:'Chủ đề Mạng xã hội & Truyền thông - Giao tiếp và ứng xử online', time:90 },
    { id:21, title:'Đề thi thử số 21', desc:'Chủ đề Lịch sử & Di tích - Đọc hiểu về di sản văn hóa Việt Nam', time:90 },
    { id:22, title:'Đề thi thử số 22', desc:'Chủ đề Thiên nhiên & Động vật - Từ vựng về bảo tồn và sinh thái', time:90 },
    { id:23, title:'Đề thi thử số 23', desc:'Chủ đề Sách & Văn hóa đọc - Đọc hiểu về thói quen đọc sách', time:90 },
    { id:24, title:'Đề thi thử số 24', desc:'Chủ đề Năng lượng tái tạo - Câu điều kiện và câu bị động', time:90 },
    { id:25, title:'Đề thi thử số 25', desc:'Chủ đề Đô thị thông minh - Từ vựng về thành phố tương lai', time:90 }
  ];
  tests.forEach(function(t){ testsSheet.appendRow([t.id,t.title,t.desc,t.time]); });

  // Helper
  function addQ(sheet, testId, questions) {
    questions.forEach(function(q, i) {
      var pn = (i < 4) ? 1 : (i < 16) ? 2 : (i < 28) ? 3 : 4;
      sheet.appendRow(['T'+testId+'_Q'+(i+1), testId, pn, q[0], q[1], q[2]||'', q[3]||'', q[4]||'', q[5]||'', q[6]||'', q[7], q[8]||0.25, q[9]||'']);
    });
  }

  // ==================== TEST 01: De minh hoa 2025 ====================
  addQ(questionsSheet, 1, [
    ['Pronunciation: -ed endings','A. watched B. stopped C. needed D. washed','','A. watched /t/','B. stopped /t/','C. needed /ɪd/','D. washed /t/','C',0.25],
    ['Pronunciation: -s/-es endings','A. books B. pens C. maps D. cats','','A. books /s/','B. pens /z/','C. maps /s/','D. cats /s/','B',0.25],
    ['Word Stress: 2-syllable','A. father B. mother C. believe D. brother','','A. father(am1)','B. mother(am1)','C. believe(am2)','D. brother(am1)','C',0.25],
    ['Word Stress: 3-syllable','A. family B. computer C. holiday D. cinema','','A. family(am1)','B. computer(am2)','C. holiday(am1)','D. cinema(am1)','B',0.25],
    ['Vocabulary in Context','The government is trying to ____ people to use public buses. A. complain B. provide C. reduce D. persuade','','A. complain','B. provide','C. reduce','D. persuade','D',0.25],
    ['Grammar: Tenses','She ____ English for five years before she moved to London. A. has studied B. had studied C. was studying D. studied','','A. has studied','B. had studied','C. was studying','D. studied','B',0.25],
    ['Phrasal Verbs','The meeting was ____ because the manager was ill. A. put off B. taken off C. given up D. turned down','','A. put off','B. taken off','C. given up','D. turned down','A',0.25],
    ['Prepositions','Valentine\'s Day is celebrated ____ February 14th. A. for B. to C. on D. at','','A. for','B. to','C. on','D. at','C',0.25],
    ['Grammar: Wish','I wish I ____ enough money to buy a new laptop. A. have B. had C. has D. having','','A. have','B. had','C. has','D. having','B',0.25],
    ['Grammar: Conditional','If I ____ you, I would accept the job offer. A. am B. was C. were D. be','','A. am','B. was','C. were','D. be','C',0.25],
    ['Vocabulary: Word Choice','Watching TV all day is a bad habit ____ we get no exercise. A. but B. even though C. because D. so','','A. but','B. even though','C. because','D. so','C',0.25],
    ['Communication','Student A: "I\'ve passed the entrance exam!" Student B: "____" A. Good idea B. Congratulations! C. Thank you D. Not at all','','A. Good idea','B. Congratulations!','C. Thank you','D. Not at all','B',0.25],
    ['Communication','Student: "Would you mind if I used your phone?" Teacher: "____" A. Yes, I\'d love to B. Of course not C. You\'re welcome D. Never mind','','A. Yes, I\'d love to','B. Of course not','C. You\'re welcome','D. Never mind','B',0.25],
    ['Public Signs','Sign: crossed-out cigarette → A. You can smoke B. Smoking allowed C. Smoking prohibited D. Cigarettes sold','','A. You can smoke','B. Smoking allowed','C. Smoking prohibited','D. Cigarettes sold','C',0.25],
    ['Public Signs','Sign: person putting trash in bin → A. Do not litter B. Trash can only C. Put trash in the bin D. Trash collection','','A. Do not litter','B. Trash can only','C. Put trash in the bin','D. Trash collection','C',0.25],
    ['Grammar: Gerund','Would you mind ____ the window? It\'s too hot. A. open B. to open C. opening D. opened','','A. open','B. to open','C. opening','D. opened','C',0.25],
    ['Cloze Test','Da Nang is one of the most beautiful cities in Vietnam. It is famous (17)____ its stunning beaches. Every year, millions (18)____ tourists visit. The city is known (19)____ its Dragon Bridge. Ba Na Hills, (20)____ is 40 km from the city, is popular. Visitors enjoy views (21)____ the mountain top. Local cuisine is a highlight, (22)____ Mi Quang and Banh Xeo.','','A. for','B. about','C. with','D. at','A',0.25],
    ['Cloze Test','Fill (18): "millions (18)____ tourists"','','A. of','B. from','C. with','D. for','A',0.25],
    ['Cloze Test','Fill (19): "known (19)____ its Dragon Bridge"','','A. of','B. about','C. for','D. with','C',0.25],
    ['Cloze Test','Fill (20): "Ba Na Hills, (20)____ is 40 km"','','A. who','B. which','C. where','D. when','B',0.25],
    ['Cloze Test','Fill (21): "views (21)____ the mountain top"','','A. in','B. on','C. at','D. from','C',0.25],
    ['Cloze Test','Fill (22): "a highlight, (22)____ Mi Quang"','','A. include','B. includes','C. included','D. including','D',0.25],
    ['Reading: MCQ','Fossil fuels are: A. unlimited & clean B. limited & harmful C. expensive & rare D. renewable & cheap','','A. unlimited clean','B. limited harmful','C. expensive rare','D. renewable cheap','B',0.25],
    ['Reading: True/False','Turning off lights saves energy. A. TRUE B. FALSE','','A. TRUE','B. FALSE','','','A',0.25],
    ['Reading: True/False','Solar and wind power are fossil fuels. A. TRUE B. FALSE','','A. TRUE','B. FALSE','','','B',0.25],
    ['Reading: True/False','Only governments are responsible for saving energy. A. TRUE B. FALSE','','A. TRUE','B. FALSE','','','B',0.25],
    ['Reading: MCQ','"Appliances" means: A. Food B. Electronic devices C. Clothes D. Cars','','A. Food','B. Electronic devices','C. Clothes','D. Cars','B',0.25],
    ['Reading: Main Idea','Main idea? A. Fossil fuels best B. Everyone can help save energy C. Only governments care D. Energy unlimited','','A. Fossil fuels best','B. Everyone can save energy','C. Only governments','D. Energy unlimited','B',0.25],
    ['Word Form','Good ____ for the final exam is essential. (PREPARE)','','','','','','preparation',0.25],
    ['Word Form','Mr. Brown is an ____ teacher. (EXPERIENCE)','','','','','','experienced',0.25],
    ['Word Form','They want to ____ the park with flowers. (BEAUTY)','','','','','','beautify',0.25],
    ['Word Form','This medicine is ____ for headaches. (EFFECT)','','','','','','effective',0.25],
    ['Word Form','Thank you for your ____. (ASSIST)','','','','','','assistance',0.25],
    ['Word Form','The team completed the project ____. (SUCCESS)','','','','','','successfully',0.25],
    ['Dictionary Insertion','"take /teɪk/ v: to move; to accept." Please ____ care of your sister.','','','','','','take',0.25],
    ['Dictionary Insertion','"follow /ˈfɒləʊ/ v: to obey or act according to." Children should ____ the example set by parents.','','','','','','follow',0.25],
    ['Sentence Transformation','Children like making models. → Children are keen ____ models.','','','','','','on making',0.25],
    ['Sentence Transformation','I advise you to see a doctor. → If I were ____ a doctor.','','','','','','you, I would see',0.25],
    ['Sentence Transformation','She last participated 2 years ago. → It has been 2 years ____.','','','','','','since she participated',0.25],
    ['Sentence Transformation','He performed well so he was nominated. → Because ____ nominated.','','','','','','he performed well, he was',0.25]
  ]);

  Logger.log('Test 1 seeded. Run seedTests02_10() next.');
}
