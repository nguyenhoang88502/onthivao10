/**
 * ============================================================
 * HELPER FUNCTIONS - Cần có trong cùng file .gs
 * ============================================================
 */

/**
 * Tạo object câu hỏi trắc nghiệm (MCQ).
 * part_number được tự động xác định dựa trên number:
 *   1-4  → Part 1 (Phonetics)
 *   5-16 → Part 2 (Grammar, Vocabulary, Communication, Signs)
 *   17-28 → Part 3 (Reading: Cloze Test + Reading Comprehension)
 */
function mcq(number, section, text, a, b, c, d, answer, explanation) {
  var partNumber;
  if (number >= 1 && number <= 4) {
    partNumber = 1;
  } else if (number >= 5 && number <= 16) {
    partNumber = 2;
  } else {
    partNumber = 3;
  }

  return {
    number: number,
    part_number: partNumber,
    section: section,
    text: text,
    option_a: a,
    option_b: b,
    option_c: c,
    option_d: d,
    answer: answer,
    explanation: explanation
  };
}

/**
 * Tạo object câu hỏi tự luận / điền từ (Text / Word Form / Writing).
 * part_number mặc định là 4 (Writing).
 */
function textQ(number, section, text, answer, explanation) {
  return {
    number: number,
    part_number: 4,
    section: section,
    text: text,
    answer: answer,
    explanation: explanation
  };
}

/**
 * Ghi tất cả câu hỏi của một đề thi vào bảng Questions.
 * @param {Object} testInfo - { test_id, title, description, time_limit, default_points }
 * @param {Array} questions - Mảng các object câu hỏi từ mcq() và textQ()
 */
function seedSingleTest_(testInfo, questions) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_QUESTIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_QUESTIONS);
    sheet.appendRow(['question_id', 'test_id', 'part_number', 'section_title',
      'question_text', 'context_image_url', 'option_a', 'option_b',
      'option_c', 'option_d', 'correct_answer', 'points', 'explanation_template']);
  }

  var rows = [];
  var testId = String(testInfo.test_id);
  var points = testInfo.default_points || CONFIG.POINTS_PER_QUESTION || 0.25;

  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    var questionId = 'T' + testId + '_Q' + q.number;

    if (q.hasOwnProperty('option_a')) {
      // MCQ question
      rows.push([
        questionId,
        testId,
        q.part_number,
        q.section,
        q.text,
        '',              // context_image_url
        q.option_a,
        q.option_b,
        q.option_c,
        q.option_d,
        q.answer,
        points,
        q.explanation
      ]);
    } else {
      // Text question
      rows.push([
        questionId,
        testId,
        q.part_number,
        q.section,
        q.text,
        '',              // context_image_url
        '',              // option_a
        '',              // option_b
        '',              // option_c
        '',              // option_d
        q.answer,
        points,
        q.explanation
      ]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 13).setValues(rows);
  }

  // Cập nhật bảng Tests nếu test_id chưa tồn tại
  var testsSheet = ss.getSheetByName(SHEET_TESTS);
  if (!testsSheet) {
    testsSheet = ss.insertSheet(SHEET_TESTS);
    testsSheet.appendRow(['test_id', 'title', 'description', 'time_limit']);
  }

  var testsData = testsSheet.getDataRange().getValues();
  var testExists = false;
  for (var t = 0; t < testsData.length; t++) {
    if (String(testsData[t][0]) === testId) {
      testExists = true;
      break;
    }
  }

  if (!testExists) {
    testsSheet.appendRow([
      testId,
      testInfo.title || ('De thi thu so ' + testId),
      testInfo.description || '',
      testInfo.time_limit || CONFIG.TIME_LIMIT_MINUTES || 90
    ]);
  }
}

// ============================================================
// HÀM TỔNG
// ============================================================

/**
 * Hàm tổng - Gọi lần lượt tất cả các hàm seed cho Đề 11 đến Đề 30.
 * Chạy hàm này một lần duy nhất để nạp toàn bộ câu hỏi mẫu vào bảng Questions.
 */
function seedAllMockQuestions() {
  seedQuestionsTest11_();
  seedQuestionsTest12_();
  seedQuestionsTest13_();
  seedQuestionsTest14_();
  seedQuestionsTest15_();
  seedQuestionsTest16_();
  seedQuestionsTest17_();
  seedQuestionsTest18_();
  seedQuestionsTest19_();
  seedQuestionsTest20_();
  seedQuestionsTest21_();
  seedQuestionsTest22_();
  seedQuestionsTest23_();
  seedQuestionsTest24_();
  seedQuestionsTest25_();
  seedQuestionsTest26_();
  seedQuestionsTest27_();
  seedQuestionsTest28_();
  seedQuestionsTest29_();
  seedQuestionsTest30_();
}

// ======================================================================
// ĐỀ 11: Chủ đề Thể thao & Thành tích - Đọc hiểu về Olympic và thể thao
// ======================================================================
function seedQuestionsTest11_() {
  var clozePassage = [
    "The Olympic Games are the world's (17)____ sports competition, held every four years.",
    "Athletes from more than 200 nations (18)____ in over 400 events across 35 different sports.",
    "The modern Olympics were revived in 1896 by Pierre de Coubertin, who was inspired (19)____ the ancient Greek tradition.",
    "Winning an Olympic medal (20)____ one of the highest honors in sport. The Games also promote values",
    "such as excellence, friendship, and respect. In recent years, the Paralympics, (21)____ takes place after",
    "the main Games, has grown in popularity. Vietnam first participated in the Olympics (22)____ 1952."
  ].join('\n\n');

  var readingPassage = [
    "Nguyen Thi Anh Vien is one of Vietnam's most celebrated athletes. Born in 1996 in Can Tho, she started",
    "swimming at the age of five. By the time she was a teenager, she had already broken several national records.",
    "At the Southeast Asian (SEA) Games, she won numerous gold medals and was nicknamed the 'Little Mermaid'",
    "by Vietnamese sports fans. Her success was not just natural talent; it was the result of years of hard work",
    "and discipline. She trained for hours every day, often waking up at 4 a.m. to practice. Besides her athletic",
    "achievements, Anh Vien also focused on her education, earning a university degree while competing internationally.",
    "Her story continues to inspire young Vietnamese athletes to pursue their dreams in sports."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. trained B. competed C. watched D. played', 'A. trained /d/', 'B. competed /ɪd/', 'C. watched /t/', 'D. played /d/', 'C', 'Âm cuối /tʃ/ của "watch" khiến -ed phát âm là /t/, khác với các từ còn lại.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. athlete B. marathon C. championship D. compete', 'A. athlete /æ/', 'B. marathon /æ/', 'C. championship /æ/', 'D. compete /ə/', 'D', '"Compete" có âm /ə/ ở âm tiết đầu, các từ còn lại có âm /æ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. athlete B. compete C. record D. tennis', 'A. athlete', 'B. compete', 'C. record', 'D. tennis', 'B', '"Compete" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. excellent B. champion C. athletics D. victory', 'A. excellent', 'B. champion', 'C. athletics', 'D. victory', 'C', '"Athletics" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'To ____ in a marathon, athletes must train for many months to build their endurance.', 'A. compete', 'B. fight', 'C. struggle', 'D. oppose', 'A', '"Compete in" là cụm từ chuẩn khi nói về tham gia thi đấu thể thao.'),
    mcq(6, 'Grammar: Tenses', 'She ____ swimming competitively since she was eight years old.', 'A. has been', 'B. was', 'C. is', 'D. had been', 'A', 'Thì hiện tại hoàn thành tiếp diễn với "since" + mốc thời gian trong quá khứ.'),
    mcq(7, 'Phrasal Verbs', 'The football match was ____ due to heavy rain and thunderstorms.', 'A. put off', 'B. taken up', 'C. given away', 'D. turned over', 'A', '"Put off" = hoãn lại, phù hợp ngữ cảnh trận đấu bị hoãn vì thời tiết xấu.'),
    mcq(8, 'Prepositions', 'Michael Phelps is famous ____ winning 23 Olympic gold medals in swimming.', 'A. for', 'B. about', 'C. with', 'D. at', 'A', 'Cấu trúc cố định: "be famous for + N/V-ing" = nổi tiếng về điều gì.'),
    mcq(9, 'Grammar: Conditional', 'If I ____ taller, I would try out for the national basketball team.', 'A. am', 'B. was', 'C. were', 'D. be', 'C', 'Câu điều kiện loại 2: If + S + were, S + would + V. Dùng "were" cho mọi ngôi.'),
    mcq(10, 'Grammar: Relative Clause', 'The stadium ____ the final match was held has a capacity of 80,000 people.', 'A. which', 'B. where', 'C. who', 'D. when', 'B', '"Where" là trạng từ quan hệ thay cho cụm chỉ nơi chốn "the stadium".'),
    mcq(11, 'Vocabulary: Word Choice', 'Regular exercise helps to ____ stress and improve mental health.', 'A. increase', 'B. reduce', 'C. produce', 'D. introduce', 'B', '"Reduce stress" (giảm căng thẳng) là cụm từ đúng trong ngữ cảnh lợi ích của thể thao.'),
    mcq(12, 'Communication', 'Nam: "I just won first prize in the school swimming competition!"\nMai: "____"', 'A. That is a good idea.', 'B. Congratulations!', 'C. Thank you very much.', 'D. Not at all.', 'B', '"Congratulations!" là lời chúc mừng phù hợp khi ai đó thông báo tin vui về thành tích.'),
    mcq(13, 'Communication', 'Coach: "Would you like to join the school basketball team?"\nStudent: "____"', 'A. Yes, I would love to!', 'B. No, I do not go.', 'C. You are welcome.', 'D. Never mind.', 'A', '"Yes, I would love to!" là lời đáp lịch sự nhận lời mời tham gia đội bóng rổ.'),
    mcq(14, 'Public Signs', 'What does this sign mean? (A sign showing a swimmer with a red cross line through it)', 'A. Swimming is allowed here', 'B. No swimming', 'C. Swimming competition ahead', 'D. Swimming pool entrance', 'B', 'Biển báo có dấu gạch chéo đỏ qua hình người bơi nghĩa là cấm bơi.'),
    mcq(15, 'Public Signs', 'A sign showing a person running toward an exit door means:', 'A. Running track ahead', 'B. Marathon start line', 'C. Emergency exit', 'D. Running is encouraged', 'C', 'Biển báo người chạy về phía cửa là biển chỉ dẫn lối thoát hiểm khẩn cấp.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She avoided ____ junk food while training for the upcoming sports event.', 'A. eat', 'B. to eat', 'C. eating', 'D. eaten', 'C', '"Avoid + V-ing" là cấu trúc cố định, nghĩa là tránh làm việc gì.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: The Olympic Games are the world\'s (17)____ sports competition.', 'A. big', 'B. bigger', 'C. biggest', 'D. most big', 'C', 'So sánh nhất với tính từ ngắn: the + adj-est. "Big" -> "biggest".'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "Athletes from more than 200 nations (18)____ in over 400 events..."', 'A. compete', 'B. competes', 'C. competing', 'D. competed', 'A', 'Chủ ngữ "Athletes" số nhiều, thì hiện tại đơn, động từ ở dạng nguyên thể.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...who was inspired (19)____ the ancient Greek tradition."', 'A. by', 'B. with', 'C. from', 'D. at', 'A', 'Câu bị động: "be inspired by" = được truyền cảm hứng bởi.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "Winning an Olympic medal (20)____ one of the highest honors..."', 'A. consider', 'B. considers', 'C. is considered', 'D. has considered', 'C', 'Câu bị động thì hiện tại đơn: S (số ít) + is + V3/ed.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...the Paralympics, (21)____ takes place after the main Games..."', 'A. who', 'B. which', 'C. where', 'D. when', 'B', 'Đại từ quan hệ "which" thay thế cho danh từ chỉ vật "the Paralympics".'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "Vietnam first participated in the Olympics (22)____ 1952."', 'A. in', 'B. on', 'C. at', 'D. since', 'A', 'Giới từ "in" dùng trước năm: in + năm.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: What is the passage mainly about?', 'A. The history of the SEA Games', 'B. The life and achievements of swimmer Nguyen Thi Anh Vien', 'C. How to become a professional swimmer', 'D. Vietnamese sports in general', 'B', 'Đoạn văn kể về tiểu sử và thành tích của vận động viên bơi lội Nguyễn Thị Ánh Viên.'),
    mcq(24, 'Reading: True/False', 'Q24: Nguyen Thi Anh Vien started swimming when she was a teenager.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu cô bắt đầu bơi lúc 5 tuổi, không phải khi là thiếu niên.'),
    mcq(25, 'Reading: True/False', 'Q25: She was nicknamed the "Little Mermaid" by Vietnamese sports fans.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu rõ biệt danh "Little Mermaid" (Nàng tiên cá nhỏ).'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, her success was mainly due to natural talent alone.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nói thành công của cô là kết quả của "years of hard work and discipline", không chỉ tài năng.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "numerous" in the passage is closest in meaning to:', 'A. few', 'B. many', 'C. rare', 'D. limited', 'B', '"Numerous" đồng nghĩa với "many" (nhiều). Cô đã giành nhiều huy chương vàng.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What message does the author want to convey through Anh Vien\'s story?', 'A. Swimming is the easiest sport to learn', 'B. Hard work and discipline lead to success', 'C. Only talented people can become athletes', 'D. University education is unimportant for athletes', 'B', 'Câu chuyện nhấn mạnh thành công đến từ sự chăm chỉ và kỷ luật, truyền cảm hứng cho giới trẻ.'),
    textQ(29, 'Word Form', 'His ____ in the 100-meter race surprised all the coaches. (PERFORM)', 'performance', 'Cần danh từ sau tính từ sở hữu "His". "Perform" (động từ) -> "performance" (danh từ).'),
    textQ(30, 'Word Form', 'The opening ceremony was ____ designed to impress the audience. (PROFESSION)', 'professionally', 'Cần trạng từ bổ nghĩa cho động từ "designed". "Profession" -> "professional" -> "professionally".'),
    textQ(31, 'Word Form', 'She is one of the most ____ athletes in the history of Vietnamese sports. (SUCCEED)', 'successful', 'Cần tính từ trong cấu trúc so sánh nhất "the most + adj". "Succeed" -> "successful".'),
    textQ(32, 'Word Form', 'The ____ of the new sports complex will begin next month. (CONSTRUCT)', 'construction', 'Cần danh từ sau mạo từ "The". "Construct" (động từ) -> "construction" (danh từ).'),
    textQ(33, 'Word Form', 'All athletes must follow a ____ diet to maintain their fitness. (HEALTH)', 'healthy', 'Cần tính từ bổ nghĩa cho danh từ "diet". "Health" (danh từ) -> "healthy" (tính từ).'),
    textQ(34, 'Word Form', 'The team celebrated their victory ____ after winning the championship. (JOY)', 'joyfully', 'Cần trạng từ bổ nghĩa cho động từ "celebrated". "Joy" -> "joyful" -> "joyfully".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "train /treɪn/ verb: to prepare for a sports event by exercising and practicing." Complete: Athletes ____ hard every day to prepare for the tournament.', 'train', 'Động từ "train" phù hợp với định nghĩa về việc tập luyện cho sự kiện thể thao.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "defeat /dɪˈfiːt/ verb: to win against someone in a fight, war, or competition." Complete: Our team managed to ____ the defending champions in the final match.', 'defeat', '"Defeat" phù hợp với ngữ cảnh đánh bại đương kim vô địch trong trận chung kết.'),
    textQ(37, 'Sentence Transformation', 'He started playing basketball five years ago. (HAS)\n→ He ____ for five years.', 'has been playing basketball', 'S + started + V-ing + time ago → S + have/has been V-ing + for + time.'),
    textQ(38, 'Sentence Transformation', 'No other player on the team is as fast as Nam. (FASTEST)\n→ Nam is ____ on the team.', 'the fastest player', 'So sánh nhất: Nam là người chạy nhanh nhất trong đội.'),
    textQ(39, 'Sentence Transformation', 'Although it rained heavily, the football match continued. (HOWEVER)\n→ It rained heavily. ____.', 'However, the football match continued', '"Although + clause" → "Clause. However, + clause" để diễn tả sự tương phản.'),
    textQ(40, 'Sentence Transformation', 'The coach said to the team, "Practice harder for the next match." (ADVISED)\n→ The coach ____ harder for the next match.', 'advised the team to practice', 'Câu tường thuật với "advise": S + advised + O + to V.'),
  ];

  seedSingleTest_({
    test_id: '11',
    title: 'Đề thi thử số 11',
    description: 'Chủ đề Thể thao & Thành tích - Đọc hiểu về Olympic và thể thao',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 12: Chủ đề Du lịch & Khám phá - Từ vựng về địa danh và hành trình
// ======================================================================
function seedQuestionsTest12_() {
  var clozePassage = [
    "Vietnam has become one of Southeast Asia's most popular travel (17)____ in recent years.",
    "From the limestone karsts of Ha Long Bay to the ancient streets of Hoi An, the country (18)____ diverse experiences.",
    "In 2024, Vietnam welcomed (19)____ 17 million international visitors, a significant increase from previous years.",
    "Tourists are drawn (20)____ the country's rich culture, delicious cuisine, and breathtaking natural landscapes.",
    "The government has also invested heavily (21)____ tourism infrastructure. Popular activities include trekking",
    "in Sapa, cruising in Ha Long Bay, and exploring the Cu Chi Tunnels. (22)____ you are an adventurous traveler,"
    + " Vietnam has something special to offer you."
  ].join('\n\n');

  var readingPassage = [
    "Ha Long Bay, located in Quang Ninh province, is one of Vietnam's most famous UNESCO World Heritage Sites.",
    "The bay features thousands of limestone islands and islets rising from emerald-green waters, creating a spectacular",
    "seascape. Legend has it that a dragon descended from the mountains, carving the landscape with its tail before",
    "plunging into the sea — hence the name 'Ha Long', meaning 'Descending Dragon'. Visitors can take boat cruises",
    "to explore caves, floating fishing villages, and hidden lagoons. The bay is also home to diverse marine life,",
    "including coral reefs and many species of fish. In recent years, authorities have implemented measures to protect",
    "the bay's environment from the negative impacts of mass tourism, ensuring that this natural wonder is preserved",
    "for future generations to enjoy."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. traveled B. visited C. explored D. enjoyed', 'A. traveled /d/', 'B. visited /ɪd/', 'C. explored /d/', 'D. enjoyed /d/', 'B', '"Visited" có âm cuối /t/ nên -ed đọc là /ɪd/, khác với các từ còn lại đọc là /d/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. beach B. trekking C. scenery D. cultural', 'A. beach /tʃ/', 'B. trekking /k/', 'C. scenery /s/', 'D. cultural /k/', 'A', '"Beach" có âm /tʃ/, các từ còn lại có âm /k/ hoặc /s/ ở phần được gạch chân.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. tourist B. hotel C. journey D. travel', 'A. tourist', 'B. hotel', 'C. journey', 'D. travel', 'B', '"Hotel" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. vacation B. holiday C. tourism D. national', 'A. vacation', 'B. holiday', 'C. tourism', 'D. national', 'A', '"Vacation" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The travel agency offers a wide ____ of tour packages to suit every budget.', 'A. range', 'B. amount', 'C. number', 'D. plenty', 'A', '"A wide range of" là cụm cố định nghĩa là "một loạt, đa dạng các".'),
    mcq(6, 'Grammar: Tenses', 'By the time we arrived at the airport, the plane ____.', 'A. took off', 'B. had taken off', 'C. has taken off', 'D. takes off', 'B', 'Quá khứ hoàn thành diễn tả hành động xảy ra trước một hành động khác trong quá khứ.'),
    mcq(7, 'Phrasal Verbs', 'We decided to ____ at a small homestay in the mountains instead of a big hotel.', 'A. put up', 'B. stay up', 'C. give up', 'D. take up', 'A', '"Put up at + nơi" = ở tạm tại một nơi nào đó trong chuyến đi.'),
    mcq(8, 'Prepositions', 'Many tourists are keen ____ trying local street food when they visit Hanoi.', 'A. on', 'B. in', 'C. at', 'D. for', 'A', 'Cấu trúc: "be keen on + N/V-ing" = thích thú, say mê điều gì.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish I ____ more time to explore all the temples in Bagan, Myanmar.', 'A. have', 'B. had', 'C. will have', 'D. having', 'B', 'Câu ước với wish ở hiện tại: S + wish + S + V (quá khứ đơn).'),
    mcq(10, 'Grammar: Passive Voice', 'Millions of tourists ____ to Vietnam each year by its natural beauty and rich culture.', 'A. attract', 'B. are attracted', 'C. attracting', 'D. have attracted', 'B', 'Câu bị động thì hiện tại đơn: S + am/is/are + V3/ed. Chủ ngữ số nhiều nên dùng "are attracted".'),
    mcq(11, 'Vocabulary: Word Choice', 'The local guide ____ us around the ancient town and told us fascinating stories.', 'A. showed', 'B. walked', 'C. moved', 'D. carried', 'A', '"Show someone around" là cụm động từ nghĩa là dẫn ai đó đi tham quan xung quanh.'),
    mcq(12, 'Communication', 'Traveler: "Could you tell me how to get to the Old Quarter?"\nLocal: "____"', 'A. Go straight ahead and turn left at the second traffic light.', 'B. I do not want to go there.', 'C. You should not visit that place.', 'D. The weather is nice today.', 'A', 'Câu trả lời chỉ đường phù hợp với câu hỏi về cách đi đến Phố Cổ.'),
    mcq(13, 'Communication', 'Guest: "I would like to book a double room for two nights."\nReceptionist: "____"', 'A. Certainly, may I have your name please?', 'B. The restaurant is on the second floor.', 'C. Check-out time is at 12 noon.', 'D. We do not have any rooms.', 'A', 'Phản hồi lịch sự của lễ tân khi khách muốn đặt phòng, hỏi thông tin để làm thủ tục.'),
    mcq(14, 'Public Signs', 'A sign at the airport showing a camera with a red cross line through it means:', 'A. Photography is allowed', 'B. No photography', 'C. Camera shop ahead', 'D. Security checkpoint', 'B', 'Biển báo có dấu gạch chéo qua hình máy ảnh nghĩa là cấm chụp ảnh.'),
    mcq(15, 'Public Signs', 'What does a sign showing a suitcase with a question mark mean?', 'A. Exit', 'B. Lost and Found', 'C. Baggage claim', 'D. Check-in counter', 'B', 'Biển báo vali có dấu hỏi là biểu tượng cho quầy Lost and Found (thất lạc hành lý).'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She suggested ____ to Da Lat for our summer vacation this year.', 'A. go', 'B. to go', 'C. going', 'D. went', 'C', '"Suggest + V-ing" là cấu trúc cố định. Không dùng "suggest + to V".'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Vietnam has become one of Southeast Asia\'s most popular travel (17)____..."', 'A. destinations', 'B. destination', 'C. places', 'D. countries', 'A', '"One of + danh từ số nhiều". "Travel destinations" = các điểm đến du lịch.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...the country (18)____ diverse experiences."', 'A. offers', 'B. offers to', 'C. is offering', 'D. has offering', 'A', 'Thì hiện tại đơn, chủ ngữ "the country" số ít nên động từ thêm -s: "offers".'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "Vietnam welcomed (19)____ 17 million international visitors..."', 'A. over', 'B. more', 'C. above', 'D. than', 'A', '"Over + con số" = hơn, trên. "Welcomed over 17 million" = đón hơn 17 triệu.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "Tourists are drawn (20)____ the country\'s rich culture..."', 'A. for', 'B. to', 'C. with', 'D. by', 'B', '"Be drawn to + N" = bị thu hút đến điều gì.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "The government has also invested heavily (21)____ tourism infrastructure."', 'A. in', 'B. on', 'C. for', 'D. at', 'A', '"Invest in + N" = đầu tư vào lĩnh vực gì.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "(22)____ you are an adventurous traveler, Vietnam has something special..."', 'A. If', 'B. Although', 'C. Because', 'D. Unless', 'A', '"If" mở đầu câu điều kiện loại 1, phù hợp ngữ cảnh.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: Where is Ha Long Bay located?', 'A. In Hanoi', 'B. In Quang Ninh province', 'C. In Hai Phong', 'D. In Ninh Binh', 'B', 'Câu đầu tiên nêu rõ: "located in Quang Ninh province".'),
    mcq(24, 'Reading: True/False', 'Q24: The name "Ha Long" means "Rising Dragon".', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu tên "Ha Long" nghĩa là "Descending Dragon" (Rồng hạ), không phải "Rising Dragon".'),
    mcq(25, 'Reading: True/False', 'Q25: Tourists can visit caves and floating villages in Ha Long Bay.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu du khách có thể khám phá "caves, floating fishing villages, and hidden lagoons".'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, mass tourism has no negative impact on Ha Long Bay.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu chính quyền đã thực hiện các biện pháp bảo vệ môi trường vịnh khỏi "negative impacts of mass tourism".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "spectacular" in the passage is closest in meaning to:', 'A. ordinary', 'B. impressive', 'C. dangerous', 'D. mysterious', 'B', '"Spectacular" = ngoạn mục, ấn tượng, đồng nghĩa với "impressive".'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main idea of this passage?', 'A. How to get to Ha Long Bay', 'B. Ha Long Bay is a natural wonder worth preserving', 'C. The history of boat cruises in Vietnam', 'D. The legend of dragons in Vietnamese culture', 'B', 'Đoạn văn mô tả vẻ đẹp của Vịnh Hạ Long và nhấn mạnh nỗ lực bảo tồn kỳ quan thiên nhiên này.'),
    textQ(29, 'Word Form', 'Ha Long Bay is one of the most ____ destinations in Vietnam. (ATTRACT)', 'attractive', 'Cần tính từ trong cấu trúc "the most + adj". "Attract" -> "attractive".'),
    textQ(30, 'Word Form', 'The local people welcomed us very ____ when we visited their village. (HOSPITALITY)', 'hospitably', 'Cần trạng từ bổ nghĩa cho động từ "welcomed". "Hospitality" -> "hospitable" -> "hospitably".'),
    textQ(31, 'Word Form', 'Tourists should show ____ when visiting religious sites in foreign countries. (RESPECTFUL)', 'respect', 'Cần danh từ sau động từ "show". "Respectful" (tính từ) -> "respect" (danh từ).'),
    textQ(32, 'Word Form', 'The ____ of the tour guide made our trip even more enjoyable. (ENTHUSIASTIC)', 'enthusiasm', 'Cần danh từ sau mạo từ "The". "Enthusiastic" (tính từ) -> "enthusiasm" (danh từ).'),
    textQ(33, 'Word Form', 'Many tourists feel ____ when they first see the beauty of Ha Long Bay. (SPEECH)', 'speechless', 'Cần tính từ sau "feel". "Speech" -> "speechless" (không nói nên lời).'),
    textQ(34, 'Word Form', 'The government should ____ more tourism campaigns to attract international visitors. (PROMOTION)', 'promote', 'Cần động từ sau "should". "Promotion" (danh từ) -> "promote" (động từ).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "explore /ɪkˈsplɔːr/ verb: to travel through an unfamiliar area to learn about it." Complete: Tourists love to ____ the narrow streets of Hoi An Ancient Town.', 'explore', '"Explore" phù hợp với định nghĩa về việc khám phá những con phố của Hội An.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "depart /dɪˈpɑːrt/ verb: to leave a place, especially to start a journey." Complete: The flight to Singapore will ____ from Tan Son Nhat Airport at 10 a.m.', 'depart', '"Depart" phù hợp với ngữ cảnh chuyến bay khởi hành từ sân bay.'),
    textQ(37, 'Sentence Transformation', 'They have never visited a more beautiful place than Ha Long Bay. (MOST)\n→ Ha Long Bay is ____ they have ever visited.', 'the most beautiful place', 'So sánh nhất: Đây là nơi đẹp nhất họ từng đến. Cấu trúc: the + most + adj + N.'),
    textQ(38, 'Sentence Transformation', 'It is a good idea to book your flight tickets in advance. (SHOULD)\n→ You ____ in advance.', 'should book your flight tickets', '"It is a good idea to V" → "You should V". Diễn tả lời khuyên.'),
    textQ(39, 'Sentence Transformation', 'The tour guide said, "Do not take photos inside the temple." (ALLOWED)\n→ The tour guide said we ____ inside the temple.', 'were not allowed to take photos', '"Do not V" trong câu tường thuật → "S + were/was not allowed to V".'),
    textQ(40, 'Sentence Transformation', 'She has not traveled abroad since she started university. (LAST)\n→ The ____ was before she started university.', 'last time she traveled abroad', '"S + have/has not V3/ed + since..." → "The last time + S + V (quá khứ) + was...".'),
  ];

  seedSingleTest_({
    test_id: '12',
    title: 'Đề thi thử số 12',
    description: 'Chủ đề Du lịch & Khám phá - Từ vựng về địa danh và hành trình',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 13: Chủ đề Thời trang & Phong cách - Mệnh đề quan hệ và câu so sánh
// ======================================================================
function seedQuestionsTest13_() {
  var clozePassage = [
    "Fashion is not just about clothing; it is a form of self-expression (17)____ reflects our personality and mood.",
    "Throughout history, people have used fashion to communicate their social status, cultural (18)____, and even political",
    "views. In recent decades, however, the rise of fast fashion has raised serious concerns. Fast fashion brands produce",
    "cheap clothing at a rapid pace, (19)____ leads to massive waste and environmental damage. Many consumers are now",
    "shifting towards sustainable fashion, which focuses (20)____ ethical production and eco-friendly materials. Buying",
    "second-hand clothes and supporting local designers are simple (21)____ to reduce our fashion footprint. The key is",
    "to buy (22)____ items of higher quality rather than many cheap ones."
  ].join('\n\n');

  var readingPassage = [
    "The traditional Vietnamese ao dai is one of the most elegant and recognizable national costumes in the world.",
    "Consisting of a long, fitted tunic worn over loose trousers, the ao dai has evolved over centuries while maintaining",
    "its graceful silhouette. In the 1930s, French-trained artist Cat Tuong redesigned the ao dai, giving it the modern",
    "form we know today. For many years, the ao dai was worn daily by women across Vietnam. Today, while it is less",
    "common as everyday wear, it remains the preferred attire for formal occasions, school uniforms at many institutions,",
    "and national celebrations. International fashion designers have also been inspired by the ao dai, incorporating its",
    "elements into modern collections. The ao dai is a proud symbol of Vietnamese cultural identity and elegance."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. designed B. inspired C. produced D. considered', 'A. designed /d/', 'B. inspired /d/', 'C. produced /t/', 'D. considered /d/', 'C', '"Produced" có âm cuối /s/ vô thanh nên -ed đọc là /t/, khác với các từ còn lại đọc /d/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. fashion B. casual C. fabric D. elegant', 'A. fashion /æ/', 'B. casual /æ/', 'C. fabric /æ/', 'D. elegant /e/', 'D', '"Elegant" có âm /e/, các từ còn lại có âm /æ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. fashion B. design C. fabric D. modern', 'A. fashion', 'B. design', 'C. fabric', 'D. modern', 'B', '"Design" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. elegant B. designer C. colorful D. quality', 'A. elegant', 'B. designer', 'C. colorful', 'D. quality', 'B', '"Designer" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The fashion show ____ the latest collection of a talented young Vietnamese designer.', 'A. featured', 'B. presented', 'C. displayed', 'D. appeared', 'A', '"Feature" = giới thiệu, trình diễn nổi bật. "Featured the latest collection" là cụm phổ biến trong lĩnh vực thời trang.'),
    mcq(6, 'Grammar: Tenses', 'Vietnamese ao dai ____ significantly over the past century, adapting to modern trends.', 'A. changed', 'B. has changed', 'C. was changing', 'D. is changing', 'B', 'Thì hiện tại hoàn thành với "over the past century" diễn tả sự thay đổi từ quá khứ đến nay.'),
    mcq(7, 'Phrasal Verbs', 'She decided to ____ the job offer from the fashion magazine to focus on her own brand.', 'A. turn down', 'B. look after', 'C. put on', 'D. give up', 'A', '"Turn down" = từ chối. "Turn down the job offer" = từ chối lời mời làm việc.'),
    mcq(8, 'Prepositions', 'This silk dress was designed ____ a famous Vietnamese designer for an international fashion show.', 'A. by', 'B. from', 'C. with', 'D. of', 'A', 'Câu bị động: "be designed by + người thiết kế".'),
    mcq(9, 'Grammar: Comparison', 'This silk scarf is ____ than the cotton one I bought at the market last week.', 'A. more expensive', 'B. most expensive', 'C. expensive', 'D. as expensive', 'A', 'So sánh hơn với tính từ dài 3 âm tiết: more + adj + than.'),
    mcq(10, 'Grammar: Relative Clause', 'The designer ____ collection won the international award graduated from a local art school.', 'A. who', 'B. whose', 'C. whom', 'D. which', 'B', '"Whose" là đại từ quan hệ sở hữu, thay cho "the designer\'s".'),
    mcq(11, 'Vocabulary: Word Choice', 'Wearing uniforms helps to ____ a sense of equality among students at school.', 'A. create', 'B. do', 'C. make', 'D. have', 'A', '"Create a sense of equality" = tạo ra cảm giác bình đẳng, là cụm từ phù hợp nhất.'),
    mcq(12, 'Communication', 'Customer: "Does this jacket come in a larger size?"\nShop assistant: "____"', 'A. Let me check the stock for you.', 'B. The price is on the tag.', 'C. We close at 9 p.m.', 'D. You should buy a different color.', 'A', 'Phản hồi lịch sự của nhân viên bán hàng khi khách hỏi về size lớn hơn.'),
    mcq(13, 'Communication', 'Friend A: "What do you think of my new dress?"\nFriend B: "____"', 'A. It looks great on you!', 'B. I bought it yesterday.', 'C. The store is far from here.', 'D. Do you have a receipt?', 'A', 'Lời khen phù hợp khi được hỏi ý kiến về trang phục mới.'),
    mcq(14, 'Public Signs', 'A sign showing a hanger with clothes and the text "Try Before You Buy" means:', 'A. No entry', 'B. Fitting room', 'C. Dry cleaning service', 'D. Clothes donation box', 'B', 'Biển báo kèm chữ "Try Before You Buy" chỉ phòng thử đồ (fitting room).'),
    mcq(15, 'Public Signs', 'What does a sign with a crossed-out shopping cart at a store entrance mean?', 'A. Free shopping carts available', 'B. Do not take shopping carts outside', 'C. Shopping carts must be returned here', 'D. Sale section', 'B', 'Biển báo gạch chéo giỏ hàng ở lối ra vào có nghĩa là không được mang giỏ hàng ra ngoài.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She cannot afford ____ designer clothes, so she shops at thrift stores instead.', 'A. buy', 'B. to buy', 'C. buying', 'D. bought', 'B', '"Afford + to V" là cấu trúc cố định, nghĩa là có đủ khả năng tài chính để làm gì.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "...a form of self-expression (17)____ reflects our personality and mood."', 'A. who', 'B. which', 'C. what', 'D. where', 'B', '"Which" thay cho "a form of self-expression" (chỉ vật) làm chủ ngữ trong mệnh đề quan hệ.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...people have used fashion to communicate their social status, cultural (18)____..."', 'A. identity', 'B. identical', 'C. identify', 'D. identification', 'A', '"Cultural identity" = bản sắc văn hóa, là cụm danh từ cố định.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...which produces cheap clothing at a rapid pace, (19)____ leads to massive waste..."', 'A. who', 'B. that', 'C. which', 'D. what', 'C', '"Which" thay thế cho cả mệnh đề phía trước, trong mệnh đề quan hệ không xác định.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...sustainable fashion, which focuses (20)____ ethical production..."', 'A. on', 'B. in', 'C. at', 'D. to', 'A', '"Focus on + N/V-ing" = tập trung vào điều gì.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "Buying second-hand clothes...are simple (21)____ to reduce our fashion footprint."', 'A. ways', 'B. methods', 'C. solutions', 'D. steps', 'A', '"Simple ways to + V" = những cách đơn giản để làm gì, phù hợp nhất về ngữ nghĩa.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "The key is to buy (22)____ items of higher quality rather than many cheap ones."', 'A. fewer', 'B. less', 'C. little', 'D. few', 'A', '"Fewer" + danh từ đếm được số nhiều "items". So sánh: fewer...rather than many.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: What is the passage mainly about?', 'A. How to sew an ao dai', 'B. The history of French fashion in Vietnam', 'C. The ao dai as a symbol of Vietnamese culture', 'D. School uniforms around the world', 'C', 'Đoạn văn giới thiệu về áo dài Việt Nam: lịch sử, đặc điểm và vai trò như biểu tượng văn hóa.'),
    mcq(24, 'Reading: True/False', 'Q24: The ao dai consists of a short tunic and tight trousers.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu áo dài gồm "a long, fitted tunic worn over loose trousers", không phải áo ngắn và quần bó.'),
    mcq(25, 'Reading: True/False', 'Q25: Cat Tuong redesigned the ao dai in the 1930s.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "In the 1930s, French-trained artist Cat Tuong redesigned the ao dai".'),
    mcq(26, 'Reading: True/False', 'Q26: Today, the ao dai is no longer worn by anyone in Vietnam.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu áo dài vẫn được mặc trong các dịp trang trọng, đồng phục trường học, và lễ kỷ niệm quốc gia.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "attire" in the passage is closest in meaning to:', 'A. food', 'B. clothing', 'C. music', 'D. festival', 'B', '"Attire" = trang phục, quần áo. Đồng nghĩa với "clothing".'),
    mcq(28, 'Reading: Main Idea', 'Q28: What can be inferred from the passage?', 'A. The ao dai has completely lost its cultural value', 'B. The ao dai continues to influence fashion beyond Vietnam', 'C. Only elderly people wear the ao dai in Vietnam', 'D. The ao dai was invented by French designers', 'B', 'Đoạn văn nêu các nhà thiết kế quốc tế đã lấy cảm hứng từ áo dài, cho thấy tầm ảnh hưởng của nó vượt ra ngoài Việt Nam.'),
    textQ(29, 'Word Form', 'The fashion show featured designs from several ____ young designers. (TALENT)', 'talented', 'Cần tính từ bổ nghĩa cho "young designers". "Talent" -> "talented".'),
    textQ(30, 'Word Form', 'She dressed ____ for the awards ceremony and caught everyone\'s attention. (ELEGANT)', 'elegantly', 'Cần trạng từ bổ nghĩa cho động từ "dressed". "Elegant" -> "elegantly".'),
    textQ(31, 'Word Form', 'The ao dai is a ____ of Vietnamese culture and national identity. (SYMBOLIZE)', 'symbol', 'Cần danh từ sau mạo từ "a". "Symbolize" (động từ) -> "symbol" (danh từ).'),
    textQ(32, 'Word Form', 'Fast fashion contributes to ____ pollution and waste in many developing countries. (ENVIRONMENT)', 'environmental', 'Cần tính từ bổ nghĩa cho danh từ "pollution". "Environment" -> "environmental".'),
    textQ(33, 'Word Form', 'Hoi An is ____ for its skilled tailors who can make custom clothing in just a day. (FAME)', 'famous', 'Cần tính từ trong cấu trúc "be ____ for". "Fame" (danh từ) -> "famous" (tính từ).'),
    textQ(34, 'Word Form', 'The ____ between the two fashion brands has resulted in an exciting new collection. (COLLABORATE)', 'collaboration', 'Cần danh từ sau mạo từ "The". "Collaborate" (động từ) -> "collaboration" (danh từ).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "trend /trend/ noun: a general direction in which something is developing or changing." Complete: Wearing vintage clothing has become a popular ____ among young people.', 'trend', '"Trend" (xu hướng) phù hợp với định nghĩa và ngữ cảnh về thời trang vintage.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "customize /ˈkʌstəmaɪz/ verb: to modify something to suit a particular individual or task." Complete: Many tailors in Hoi An can ____ a dress to fit you perfectly within hours.', 'customize', '"Customize" (tùy chỉnh) phù hợp với ngữ cảnh may đo theo yêu cầu riêng.'),
    textQ(37, 'Sentence Transformation', 'No other fashion show in the city is as popular as this one. (MOST)\n→ This is ____ show in the city.', 'the most popular fashion', 'So sánh nhất: "No other...is as...as" → "This is the most + adj + N".'),
    textQ(38, 'Sentence Transformation', 'She started designing her own clothes three years ago. (BEEN)\n→ She ____ her own clothes for three years.', 'has been designing', '"S + started V-ing + time ago" → "S + have/has been V-ing + for + time".'),
    textQ(39, 'Sentence Transformation', '"Be careful when washing this silk blouse," the saleswoman said. (ADVISED)\n→ The saleswoman ____ when washing that silk blouse.', 'advised me to be careful', 'Câu tường thuật với "advise": S + advised + O + to V.'),
    textQ(40, 'Sentence Transformation', 'She cannot buy that dress because it is too expensive. (AFFORD)\n→ She cannot ____ that dress.', 'afford to buy', '"Cannot afford to V" = không đủ khả năng tài chính để làm gì. Dùng "afford" trực tiếp.'),
  ];

  seedSingleTest_({
    test_id: '13',
    title: 'Đề thi thử số 13',
    description: 'Chủ đề Thời trang & Phong cách - Mệnh đề quan hệ và câu so sánh',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 14: Chủ đề Nghề nghiệp & Tương lai - Ước mơ và định hướng nghề nghiệp
// ======================================================================
function seedQuestionsTest14_() {
  var clozePassage = [
    "Choosing a career path is one of the most important (17)____ a young person will make in life.",
    "Many students feel pressured to pursue traditional professions (18)____ as medicine, law, or engineering.",
    "However, the job market is constantly evolving, and new careers (19)____ every year in fields like data science,",
    "digital marketing, and renewable energy. Experts advise young people (20)____ their passions and strengths when",
    "making career decisions. It is also important to (21)____ practical experience through internships and part-time",
    "jobs during school years. The key to a fulfilling career is finding work that (22)____ your interests with market demand."
  ].join('\n\n');

  var readingPassage = [
    "In today's rapidly changing world, the skills needed for career success are very different from what they were",
    "a generation ago. While academic knowledge remains important, employers increasingly value soft skills such as",
    "communication, teamwork, problem-solving, and adaptability. A recent survey of companies in Vietnam found that",
    "85% of employers consider soft skills equally or more important than technical qualifications when hiring new",
    "graduates. Furthermore, with the rise of artificial intelligence and automation, many routine jobs are disappearing,",
    "while new roles requiring creativity and emotional intelligence are growing. This shift means that students should",
    "focus not only on getting good grades but also on developing well-rounded skills through extracurricular activities,",
    "volunteer work, and real-world projects. Learning how to learn — and how to adapt — may be the most valuable",
    "skill of all in the 21st-century workplace."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -s/-es endings', 'Choose the word whose underlined part is pronounced differently: A. careers B. skills C. jobs D. paths', 'A. careers /z/', 'B. skills /z/', 'C. jobs /z/', 'D. paths /θs/', 'D', '"Paths" có âm cuối /θ/ vô thanh nên -s đọc là /s/, các từ còn lại đọc /z/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. engineer B. career C. volunteer D. peer', 'A. engineer /ɪə/', 'B. career /ɪə/', 'C. volunteer /ɪə/', 'D. peer /ɪə/', 'D', '"Peer" là từ đơn âm tiết, trong khi các từ còn lại là từ đa âm tiết có trọng âm khác nhau.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. career B. future C. doctor D. teacher', 'A. career', 'B. future', 'C. doctor', 'D. teacher', 'A', '"Career" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. interview B. engineer C. manager D. scientist', 'A. interview', 'B. engineer', 'C. manager', 'D. scientist', 'B', '"Engineer" nhấn âm 3, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'A good employee should be ____ and willing to learn new skills throughout their career.', 'A. adaptable', 'B. comfortable', 'C. portable', 'D. preventable', 'A', '"Adaptable" (có khả năng thích nghi) là phẩm chất quan trọng trong môi trường làm việc hiện đại.'),
    mcq(6, 'Grammar: Tenses', 'By the time she graduates next year, she ____ English for twelve years.', 'A. will study', 'B. will have studied', 'C. studied', 'D. has studied', 'B', 'Thì tương lai hoàn thành: By the time + HTĐ, S + will have V3/ed.'),
    mcq(7, 'Phrasal Verbs', 'After much consideration, she decided to ____ the job offer and move to a different city.', 'A. take up', 'B. put off', 'C. give in', 'D. turn over', 'A', '"Take up a job offer" = nhận lời mời làm việc.'),
    mcq(8, 'Prepositions', 'Many young people are interested ____ working in the technology sector after graduation.', 'A. at', 'B. in', 'C. on', 'D. with', 'B', 'Cấu trúc: "be interested in + N/V-ing".'),
    mcq(9, 'Grammar: Conditional', 'If she ____ more confident during the interview, she would have gotten the job.', 'A. is', 'B. was', 'C. were', 'D. had been', 'D', 'Câu điều kiện loại 3: If + S + had V3/ed, S + would have V3/ed.'),
    mcq(10, 'Grammar: Reported Speech', 'The career counselor asked me what ____ to do after high school graduation.', 'A. I wanted', 'B. did I want', 'C. do I want', 'D. I want', 'A', 'Câu tường thuật: lùi thì từ hiện tại đơn "want" → quá khứ đơn "wanted", không đảo trợ động từ.'),
    mcq(11, 'Vocabulary: Word Choice', 'The new apprenticeship program ____ students with valuable hands-on experience in their chosen field.', 'A. provides', 'B. gives', 'C. offers', 'D. supplies', 'A', '"Provide someone with something" là cấu trúc cố định, phù hợp nhất với ngữ cảnh.'),
    mcq(12, 'Communication', 'Lan: "What do you want to be in the future?"\nNam: "____"', 'A. I dream of becoming a software engineer.', 'B. I am 15 years old.', 'C. I go to school every day.', 'D. My father is a doctor.', 'A', 'Câu trả lời về nghề nghiệp mơ ước phù hợp với câu hỏi về định hướng tương lai.'),
    mcq(13, 'Communication', 'Student: "Could you give me some advice on choosing a career?"\nTeacher: "____"', 'A. You should consider both your interests and job market demand.', 'B. I am busy right now.', 'C. The exam will be next week.', 'D. Your uniform looks nice today.', 'A', 'Lời khuyên về việc chọn nghề nghiệp phù hợp với câu hỏi xin lời khuyên.'),
    mcq(14, 'Public Signs', 'What does a sign showing a person with a hard hat mean?', 'A. Restaurant area', 'B. Construction site — hard hats required', 'C. Hospital zone', 'D. School crossing', 'B', 'Biển báo người đội mũ bảo hộ chỉ khu vực công trường, yêu cầu đội mũ cứng.'),
    mcq(15, 'Public Signs', 'A sign at a workplace showing "Fire Extinguisher" with an arrow indicates:', 'A. No smoking area', 'B. Location of fire extinguisher', 'C. Fire exit', 'D. Fire alarm button', 'B', 'Biển báo "Fire Extinguisher" kèm mũi tên chỉ vị trí đặt bình chữa cháy.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She decided ____ a gap year before starting university to gain more life experience.', 'A. take', 'B. to take', 'C. taking', 'D. took', 'B', '"Decide + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "...one of the most important (17)____ a young person will make in life."', 'A. decisions', 'B. choices', 'C. options', 'D. selections', 'A', '"Make decisions" là cụm cố định, phù hợp nhất với ngữ cảnh chọn nghề nghiệp.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...traditional professions (18)____ as medicine, law, or engineering."', 'A. such', 'B. like', 'C. example', 'D. similar', 'A', '"Such as" = ví dụ như, dùng để liệt kê các ví dụ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...new careers (19)____ every year in fields like data science..."', 'A. emerge', 'B. emerges', 'C. emerging', 'D. emerged', 'A', '"Careers" số nhiều, thì hiện tại đơn: emerge (không chia).'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "Experts advise young people (20)____ their passions and strengths..."', 'A. consider', 'B. considering', 'C. to consider', 'D. considered', 'C', '"Advise someone to V" là cấu trúc cố định.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "It is also important to (21)____ practical experience through internships..."', 'A. gain', 'B. make', 'C. do', 'D. take', 'A', '"Gain experience" = tích lũy kinh nghiệm, là cụm cố định.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...finding work that (22)____ your interests with market demand."', 'A. balances', 'B. connects', 'C. joins', 'D. links', 'A', '"Balance A with B" = cân bằng giữa sở thích và nhu cầu thị trường.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, which skills are increasingly valued by employers today?', 'A. Only technical qualifications', 'B. Soft skills such as communication and teamwork', 'C. Memorization and repetition', 'D. Physical strength and speed', 'B', 'Đoạn văn nêu: "employers increasingly value soft skills such as communication, teamwork, problem-solving, and adaptability".'),
    mcq(24, 'Reading: True/False', 'Q24: According to a survey, 85% of employers in Vietnam value soft skills as much as or more than technical qualifications.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "85% of employers consider soft skills equally or more important than technical qualifications".'),
    mcq(25, 'Reading: True/False', 'Q25: The passage suggests that AI and automation will create more routine jobs for humans.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu AI và tự động hóa khiến "routine jobs are disappearing", tức là công việc lặp lại đang biến mất.'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, getting good grades is the only thing students need for future success.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu học sinh nên phát triển "well-rounded skills" thông qua nhiều hoạt động, không chỉ tập trung vào điểm số.'),
    mcq(27, 'Reading: MCQ', 'Q27: The phrase "well-rounded skills" in the passage most likely means:', 'A. Skills related to sports only', 'B. A variety of different skills and abilities', 'C. Skills in mathematics and science', 'D. Skills that are difficult to learn', 'B', '"Well-rounded" = toàn diện, đa dạng. Chỉ nhiều kỹ năng khác nhau.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the author\'s main message about career preparation?', 'A. Only academic grades matter for career success', 'B. Students should focus on developing both academic knowledge and soft skills', 'C. Technology will replace all human jobs in the future', 'D. Changing careers is too difficult for most people', 'B', 'Bài đọc nhấn mạnh tầm quan trọng của việc phát triển cả kiến thức học thuật lẫn kỹ năng mềm.'),
    textQ(29, 'Word Form', 'Her ____ to become a doctor has inspired many young girls in her village. (DETERMINE)', 'determination', 'Cần danh từ sau tính từ sở hữu "Her". "Determine" (động từ) -> "determination" (danh từ).'),
    textQ(30, 'Word Form', 'He is a highly ____ engineer with more than 15 years in the automotive industry. (EXPERIENCE)', 'experienced', 'Cần tính từ bổ nghĩa cho "engineer". "Experience" -> "experienced" (giàu kinh nghiệm).'),
    textQ(31, 'Word Form', 'The ____ fair at our school attracted over 50 companies looking to recruit new graduates. (CAREER)', 'career', '"Career fair" = hội chợ việc làm, là danh từ ghép cố định. Giữ nguyên "career".'),
    textQ(32, 'Word Form', 'You need to show ____ and patience when dealing with difficult customers. (CREATE)', 'creativity', 'Cần danh từ song song với "patience". "Create" -> "creativity" (sự sáng tạo).'),
    textQ(33, 'Word Form', 'She prepared ____ for the interview and ended up getting her dream job. (THOROUGH)', 'thoroughly', 'Cần trạng từ bổ nghĩa cho động từ "prepared". "Thorough" -> "thoroughly".'),
    textQ(34, 'Word Form', 'The company offers many training programs for the ____ development of its employees. (PROFESSION)', 'professional', 'Cần tính từ bổ nghĩa cho danh từ "development". "Profession" -> "professional".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "apply /əˈplaɪ/ verb: to make a formal request for a job, position, or course." Complete: Students should ____ for scholarships before the deadline at the end of this month.', 'apply', '"Apply for" (nộp đơn xin) phù hợp với định nghĩa và ngữ cảnh xin học bổng.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "qualify /ˈkwɒlɪfaɪ/ verb: to have the necessary skills, knowledge, or experience to do something." Complete: After completing the training course, you will ____ to work as a certified electrician.', 'qualify', '"Qualify to V" (đủ điều kiện để làm gì) phù hợp với ngữ cảnh hoàn thành khóa đào tạo.'),
    textQ(37, 'Sentence Transformation', 'If I were you, I would choose a career that matches your passion. (ADVISED)\n→ She ____ a career that matches my passion.', 'advised me to choose', '"If I were you, I would V" → "S + advised + O + to V" trong lời khuyên.'),
    textQ(38, 'Sentence Transformation', 'He started working at this company six years ago. (BEEN)\n→ He ____ at this company for six years.', 'has been working', '"S + started V-ing + time ago" → "S + have/has been V-ing + for + time".'),
    textQ(39, 'Sentence Transformation', 'No other profession is as rewarding as teaching, in my opinion. (MOST)\n→ Teaching is ____ profession, in my opinion.', 'the most rewarding', 'So sánh nhất: "No other...is as...as" → "This is the most + adj + N".'),
    textQ(40, 'Sentence Transformation', '"Why do you want to become a nurse?" the interviewer asked her. (ASKED)\n→ The interviewer ____ to become a nurse.', 'asked her why she wanted', 'Câu tường thuật Wh-question: S + asked + O + Wh-word + S + V (lùi thì).'),
  ];

  seedSingleTest_({
    test_id: '14',
    title: 'Đề thi thử số 14',
    description: 'Chủ đề Nghề nghiệp & Tương lai - Ước mơ và định hướng nghề nghiệp',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 15: Chủ đề Tình nguyện & Cộng đồng - Đọc hiểu về hoạt động xã hội
// ======================================================================
function seedQuestionsTest15_() {
  var clozePassage = [
    "Volunteering plays a vital role (17)____ building strong and connected communities around the world.",
    "Young people (18)____ participate in volunteer activities develop empathy, leadership skills, and a sense of",
    "social responsibility. Common volunteer opportunities (19)____ helping at local shelters, organizing charity events,",
    "and tutoring underprivileged children. In Vietnam, youth volunteer movements (20)____ as the Green Summer Campaign",
    "have made significant contributions to rural development. Studies show that volunteers not only benefit their",
    "communities (21)____ also experience improved mental health and life satisfaction. Anyone can volunteer, regardless",
    "of age (22)____ background — the most important thing is the willingness to help others."
  ].join('\n\n');

  var readingPassage = [
    "The Green Summer Campaign is one of Vietnam's largest and longest-running youth volunteer programs. Organized",
    "by the Ho Chi Minh Communist Youth Union, it takes place every summer and attracts thousands of university",
    "students from across the country. During the campaign, volunteers travel to rural and remote areas to help build",
    "roads, bridges, and houses for disadvantaged families. They also teach children, provide free medical check-ups,",
    "and organize environmental clean-up activities. The campaign, which began in 1997, has had a lasting impact on",
    "countless communities. More importantly, it helps young volunteers develop a deep understanding of social issues",
    "and a lifelong commitment to community service. Many former volunteers say the experience changed their",
    "perspective on life and inspired them to pursue careers in social work, education, and public service."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. helped B. volunteered C. organized D. donated', 'A. helped /t/', 'B. volunteered /d/', 'C. organized /d/', 'D. donated /ɪd/', 'A', '"Helped" có âm cuối /p/ vô thanh nên -ed đọc /t/, khác với các từ còn lại.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. charity B. children C. school D. achieve', 'A. charity /tʃ/', 'B. children /tʃ/', 'C. school /k/', 'D. achieve /tʃ/', 'C', '"School" có âm "ch" phát âm là /k/, các từ còn lại có âm /tʃ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. volunteer B. service C. kindness D. helpful', 'A. volunteer', 'B. service', 'C. kindness', 'D. helpful', 'A', '"Volunteer" nhấn âm 3 (từ 3 âm tiết), các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. charity B. community C. organize D. benefit', 'A. charity', 'B. community', 'C. organize', 'D. benefit', 'B', '"Community" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The charity organization was set up to help ____ children access education in remote mountainous areas.', 'A. underprivileged', 'B. wealthy', 'C. generous', 'D. fortunate', 'A', '"Underprivileged" (thiệt thòi, kém may mắn) phù hợp với ngữ cảnh trẻ em vùng sâu vùng xa cần giúp đỡ.'),
    mcq(6, 'Grammar: Tenses', 'Since the program began, thousands of volunteers ____ in community service activities across the country.', 'A. participated', 'B. have participated', 'C. participate', 'D. were participating', 'B', 'Thì hiện tại hoàn thành với "Since" + mệnh đề quá khứ.'),
    mcq(7, 'Phrasal Verbs', 'We need more volunteers to help ____ food and supplies to the flood victims.', 'A. give out', 'B. take off', 'C. put away', 'D. turn down', 'A', '"Give out" = phân phát. "Give out food and supplies" = phân phát thực phẩm và nhu yếu phẩm.'),
    mcq(8, 'Prepositions', 'Many students at our school are involved ____ volunteer work during the summer holidays.', 'A. in', 'B. on', 'C. at', 'D. with', 'A', 'Cấu trúc: "be involved in + N/V-ing" = tham gia vào.'),
    mcq(9, 'Grammar: Conditional', 'If more people volunteered their time, the community ____ a much better place to live.', 'A. will be', 'B. would be', 'C. is', 'D. was', 'B', 'Câu điều kiện loại 2: If + S + V (QKĐ), S + would + V.'),
    mcq(10, 'Grammar: Wish Clause', 'I wish I ____ more free time so that I could volunteer at the local orphanage.', 'A. have', 'B. had', 'C. will have', 'D. having', 'B', 'Câu ước wish ở hiện tại: S + wish + S + V (quá khứ đơn).'),
    mcq(11, 'Vocabulary: Word Choice', 'The Green Summer Campaign has made a significant ____ to improving rural infrastructure.', 'A. contribution', 'B. donation', 'C. support', 'D. help', 'A', '"Make a contribution to + N/V-ing" = đóng góp vào. Là cụm cố định.'),
    mcq(12, 'Communication', 'Student A: "Would you like to join our charity run this weekend?"\nStudent B: "____"', 'A. I would love to! What time does it start?', 'B. I do not like running.', 'C. You are welcome.', 'D. That is not my problem.', 'A', 'Lời đáp nhận lời mời tham gia chạy từ thiện một cách lịch sự và hào hứng.'),
    mcq(13, 'Communication', 'Organizer: "Thank you so much for volunteering today."\nVolunteer: "____"', 'A. You are welcome. I am happy to help.', 'B. No, I do not want to.', 'C. I am not sure about that.', 'D. Can I go home now?', 'A', 'Phản hồi lịch sự khi được cảm ơn vì đã tham gia tình nguyện.'),
    mcq(14, 'Public Signs', 'A sign showing a hand holding a heart symbol at a community center most likely means:', 'A. Hospital entrance', 'B. Volunteer and charity point', 'C. Pharmacy', 'D. Blood pressure check', 'B', 'Biểu tượng bàn tay ôm trái tim thường chỉ điểm tình nguyện và từ thiện.'),
    mcq(15, 'Public Signs', 'What does a sign showing a person with a broom sweeping mean?', 'A. No entry', 'B. Cleaning in progress — caution wet floor', 'C. Fire exit', 'D. Emergency assembly point', 'B', 'Biển báo người đang quét dọn cảnh báo khu vực đang lau dọn, sàn có thể trơn ướt.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'They agreed ____ the fundraising event for the children\'s hospital next month.', 'A. organize', 'B. to organize', 'C. organizing', 'D. organized', 'B', '"Agree + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Volunteering plays a vital role (17)____ building strong and connected communities..."', 'A. in', 'B. on', 'C. at', 'D. for', 'A', '"Play a role in + N/V-ing" = đóng vai trò trong việc gì.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "Young people (18)____ participate in volunteer activities develop empathy..."', 'A. who', 'B. which', 'C. whom', 'D. where', 'A', '"Who" thay cho "young people" (chỉ người) làm chủ ngữ trong mệnh đề quan hệ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "Common volunteer opportunities (19)____ helping at local shelters..."', 'A. include', 'B. includes', 'C. including', 'D. included', 'A', '"Opportunities" số nhiều, thì hiện tại đơn: include (không chia).'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...youth volunteer movements (20)____ as the Green Summer Campaign..."', 'A. such', 'B. like', 'C. similar', 'D. same', 'A', '"Such as" = ví dụ như, dùng để đưa ra ví dụ cụ thể.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...volunteers not only benefit their communities (21)____ also experience..."', 'A. but', 'B. and', 'C. or', 'D. so', 'A', 'Cấu trúc "not only...but also" = không những...mà còn.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...anyone can volunteer, regardless of age (22)____ background..."', 'A. or', 'B. and', 'C. but', 'D. nor', 'A', '"Regardless of A or B" = bất kể A hay B.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: When did the Green Summer Campaign begin?', 'A. In 1995', 'B. In 1997', 'C. In 2000', 'D. In 2010', 'B', 'Đoạn văn nêu: "The campaign, which began in 1997".'),
    mcq(24, 'Reading: True/False', 'Q24: The Green Summer Campaign is organized by the Vietnamese government.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu chương trình do "Ho Chi Minh Communist Youth Union" (Đoàn Thanh niên) tổ chức, không phải chính phủ.'),
    mcq(25, 'Reading: True/False', 'Q25: During the campaign, volunteers only build houses and do not teach children.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn liệt kê nhiều hoạt động bao gồm dạy học cho trẻ em, không chỉ xây nhà.'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, the campaign has no lasting impact on volunteers themselves.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu nhiều cựu tình nguyện viên nói trải nghiệm này đã thay đổi cách nhìn cuộc sống của họ.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "commitment" in the passage is closest in meaning to:', 'A. hesitation', 'B. dedication', 'C. refusal', 'D. ignorance', 'B', '"Commitment" = sự cam kết, tận tụy. Đồng nghĩa với "dedication".'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of the Green Summer Campaign according to the passage?', 'A. To provide summer jobs for unemployed youth', 'B. To help rural communities while developing young people\'s sense of social responsibility', 'C. To recruit students for the army', 'D. To promote tourism in rural areas', 'B', 'Chiến dịch vừa giúp đỡ cộng đồng nông thôn vừa phát triển ý thức trách nhiệm xã hội cho thanh niên.'),
    textQ(29, 'Word Form', 'His ____ to the community was recognized with a special award from the city mayor. (CONTRIBUTE)', 'contribution', 'Cần danh từ sau tính từ sở hữu "His". "Contribute" (động từ) -> "contribution" (danh từ).'),
    textQ(30, 'Word Form', 'The volunteers worked ____ to repair the damaged school before the new semester began. (TIRE)', 'tirelessly', 'Cần trạng từ bổ nghĩa cho động từ "worked". "Tire" -> "tireless" -> "tirelessly" (không mệt mỏi).'),
    textQ(31, 'Word Form', 'The charity event was a ____ success, raising over 500 million VND for flood victims. (REMARK)', 'remarkable', 'Cần tính từ bổ nghĩa cho danh từ "success". "Remark" -> "remarkable" (đáng chú ý).'),
    textQ(32, 'Word Form', 'Many young people find ____ in helping others and making a difference in their community. (HAPPY)', 'happiness', 'Cần danh từ sau động từ "find". "Happy" (tính từ) -> "happiness" (danh từ).'),
    textQ(33, 'Word Form', 'She spoke ____ about her experiences as a volunteer teacher in a remote mountain village. (PASSION)', 'passionately', 'Cần trạng từ bổ nghĩa cho động từ "spoke". "Passion" -> "passionate" -> "passionately".'),
    textQ(34, 'Word Form', 'The local community showed great ____ to the volunteers from the city. (HOSPITABLE)', 'hospitality', 'Cần danh từ sau tính từ "great". "Hospitable" (tính từ) -> "hospitality" (danh từ).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "donate /dəʊˈneɪt/ verb: to give money, food, or goods to help a person or organization." Complete: Many local businesses ____ money to support the building of a new community library.', 'donated', '"Donate" (quyên góp) phù hợp với ngữ cảnh doanh nghiệp ủng hộ tiền xây thư viện. Quá khứ đơn vì sự kiện đã hoàn thành.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "participate /pɑːrˈtɪsɪpeɪt/ verb: to take part in or become involved in an activity." Complete: Students are encouraged to ____ in at least one community service activity each semester.', 'participate', '"Participate in" (tham gia vào) phù hợp với ngữ cảnh khuyến khích học sinh tham gia hoạt động cộng đồng.'),
    textQ(37, 'Sentence Transformation', 'She began volunteering at the local orphanage three years ago. (HAS)\n→ She ____ at the local orphanage for three years.', 'has been volunteering', '"S + began V-ing + time ago" → "S + have/has been V-ing + for + time".'),
    textQ(38, 'Sentence Transformation', '"Let\'s organize a fundraising event for the homeless," said the team leader. (SUGGESTED)\n→ The team leader ____ a fundraising event for the homeless.', 'suggested organizing', '"Let\'s V" trong câu tường thuật → "S + suggested + V-ing".'),
    textQ(39, 'Sentence Transformation', 'They have never participated in such a meaningful volunteer campaign before. (FIRST)\n→ This is the ____ in such a meaningful volunteer campaign.', 'first time they have participated', '"S + have/has never V3/ed...before" → "This is the first time + S + have/has V3/ed".'),
    textQ(40, 'Sentence Transformation', 'Because of her dedication to community service, she was given an award. (SO)\n→ She was ____ given an award.', 'so dedicated to community service that she was', '"Because of + N" → "so + adj + that + clause" để diễn tả nguyên nhân-kết quả.'),
  ];

  seedSingleTest_({
    test_id: '15',
    title: 'Đề thi thử số 15',
    description: 'Chủ đề Tình nguyện & Cộng đồng - Đọc hiểu về hoạt động xã hội',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 16: Bám sát đề chính thức 2024 - Đà Nẵng và tiết kiệm năng lượng
// ======================================================================
function seedQuestionsTest16_() {
  var clozePassage = [
    "Da Nang is one of the most (17)____ cities in Vietnam, located on the beautiful central coast.",
    "It is famous (18)____ its stunning beaches, the Han River, and the iconic Golden Bridge at Ba Na Hills.",
    "The city attracts millions (19)____ tourists each year from both Vietnam and abroad. In recent years,",
    "Da Nang has also become a hub (20)____ technology and innovation, hosting numerous international conferences.",
    "The local government has implemented smart city (21)____ to improve public services and reduce traffic congestion.",
    "Da Nang is also known for its delicious cuisine, (22)____ dishes like Mi Quang and Banh Trang Cuon Thit Heo."
  ].join('\n\n');

  var readingPassage = [
    "Energy conservation has become a global priority as countries work to reduce carbon emissions and combat climate",
    "change. Simple actions such as turning off lights when leaving a room, using energy-efficient appliances, and",
    "reducing water waste can make a significant difference in the long run. In Vietnam, the government has launched",
    "several campaigns encouraging households and businesses to save electricity, especially during the hot summer months",
    "when demand peaks. Renewable energy sources like solar and wind power are also being promoted as sustainable",
    "alternatives to fossil fuels. Many families have started installing solar panels on their rooftops, which not only",
    "lowers their electricity bills but also contributes to environmental protection. Experts say that if every household",
    "in Vietnam replaced just one incandescent bulb with an LED bulb, the country could save millions of kilowatt-hours",
    "of electricity each year. Saving energy is not just about reducing costs; it is about ensuring a livable planet",
    "for future generations."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. launched B. installed C. reduced D. replaced', 'A. launched /t/', 'B. installed /d/', 'C. reduced /t/', 'D. replaced /t/', 'B', '"Installed" có âm cuối /l/ hữu thanh nên -ed đọc /d/, các từ còn lại có âm cuối vô thanh /tʃ/, /s/, /s/ nên -ed đọc /t/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. energy B. electricity C. environment D. efficient', 'A. energy /e/', 'B. electricity /ɪ/', 'C. environment /ɪ/', 'D. efficient /ɪ/', 'A', '"Energy" có âm /e/ ở âm tiết đầu, các từ còn lại có âm /ɪ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. solar B. fossil C. conserve D. power', 'A. solar', 'B. fossil', 'C. conserve', 'D. power', 'C', '"Conserve" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. energy B. electric C. furniture D. holiday', 'A. energy', 'B. electric', 'C. furniture', 'D. holiday', 'B', '"Electric" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'Da Nang is ____ for its clean beaches, friendly people, and delicious local food.', 'A. famous', 'B. interested', 'C. excited', 'D. curious', 'A', '"Famous for" (nổi tiếng về) là cụm từ cố định phù hợp nhất với ngữ cảnh.'),
    mcq(6, 'Grammar: Tenses', 'Da Nang ____ significantly over the past decade, becoming a top destination in Southeast Asia.', 'A. developed', 'B. has developed', 'C. was developing', 'D. develops', 'B', 'Thì hiện tại hoàn thành với "over the past decade" diễn tả sự phát triển từ quá khứ đến nay.'),
    mcq(7, 'Phrasal Verbs', 'Please remember to ____ the lights when you leave the room to save electricity.', 'A. turn off', 'B. turn on', 'C. take off', 'D. give up', 'A', '"Turn off the lights" = tắt đèn, là hành động tiết kiệm điện cơ bản.'),
    mcq(8, 'Prepositions', 'Da Nang is located ____ the central coast of Vietnam, making it an ideal beach destination.', 'A. on', 'B. in', 'C. at', 'D. by', 'A', '"On the coast" là giới từ cố định khi nói về vị trí dọc bờ biển.'),
    mcq(9, 'Grammar: Passive Voice', 'Solar panels ____ on many rooftops in Vietnam to capture sunlight and generate electricity.', 'A. install', 'B. are installed', 'C. installed', 'D. have installed', 'B', 'Câu bị động thì hiện tại đơn: S + am/is/are + V3/ed.'),
    mcq(10, 'Grammar: Conditional', 'If every household ____ LED bulbs, the country would save a huge amount of electricity.', 'A. uses', 'B. used', 'C. will use', 'D. would use', 'B', 'Câu điều kiện loại 2: If + S + V (QKĐ), S + would + V.'),
    mcq(11, 'Vocabulary: Word Choice', 'The Golden Bridge in Ba Na Hills has become an ____ tourist attraction since its opening.', 'A. iconic', 'B. ordinary', 'C. boring', 'D. forgettable', 'A', '"Iconic" (mang tính biểu tượng) là từ phù hợp nhất mô tả Cầu Vàng nổi tiếng.'),
    mcq(12, 'Communication', 'Tourist: "Could you recommend a good local restaurant in Da Nang?"\nLocal: "____"', 'A. You should try Ba Duong for the best Banh Xeo in town.', 'B. I do not eat out very often.', 'C. The weather is very hot today.', 'D. I live far from the city center.', 'A', 'Lời giới thiệu nhà hàng địa phương phù hợp với câu hỏi xin gợi ý.'),
    mcq(13, 'Communication', 'Visitor: "Excuse me, how do I get to My Khe Beach from here?"\nPasserby: "____"', 'A. Go straight for about 500 meters, then turn right at the traffic light.', 'B. The beach is very beautiful this time of year.', 'C. I have never been to My Khe Beach.', 'D. You should bring sunscreen and a hat.', 'A', 'Chỉ đường cụ thể là câu trả lời phù hợp cho câu hỏi về cách đi đến Bãi biển Mỹ Khê.'),
    mcq(14, 'Public Signs', 'A sign showing a lightning bolt symbol on a yellow triangle means:', 'A. Wi-Fi available', 'B. Danger: High voltage electricity', 'C. Battery charging station', 'D. Power saving mode', 'B', 'Biển báo hình tam giác vàng có tia sét cảnh báo nguy hiểm điện cao thế.'),
    mcq(15, 'Public Signs', 'What does a sign with a faucet and a drop of water inside a green circle mean?', 'A. No drinking water', 'B. Water conservation area', 'C. Hot water warning', 'D. Swimming pool', 'B', 'Biển báo vòi nước với giọt nước trong vòng tròn xanh thường chỉ khu vực tiết kiệm nước.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'We should avoid ____ electricity during peak hours to reduce the load on the power grid.', 'A. waste', 'B. to waste', 'C. wasting', 'D. wasted', 'C', '"Avoid + V-ing" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Da Nang is one of the most (17)____ cities in Vietnam..."', 'A. dynamic', 'B. dynamics', 'C. dynamically', 'D. more dynamic', 'A', 'Cần tính từ trong cấu trúc "one of the most + adj". "Dynamic" (năng động) phù hợp nhất.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "It is famous (18)____ its stunning beaches, the Han River..."', 'A. for', 'B. with', 'C. about', 'D. at', 'A', '"Famous for" là cụm giới từ cố định.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "The city attracts millions (19)____ tourists each year..."', 'A. of', 'B. from', 'C. with', 'D. by', 'A', '"Millions of + N" = hàng triệu. Là cấu trúc cố định.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...a hub (20)____ technology and innovation..."', 'A. of', 'B. for', 'C. in', 'D. with', 'B', '"A hub for + N" = trung tâm cho lĩnh vực gì.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...smart city (21)____ to improve public services..."', 'A. initiatives', 'B. initiations', 'C. initials', 'D. initiates', 'A', '"Smart city initiatives" = các sáng kiến đô thị thông minh, là cụm danh từ.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...delicious cuisine, (22)____ dishes like Mi Quang..."', 'A. including', 'B. includes', 'C. included', 'D. include', 'A', '"Including" (bao gồm) là giới từ/V-ing bổ nghĩa cho danh từ phía trước.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, what is one simple way to save energy at home?', 'A. Leaving lights on all day', 'B. Using energy-efficient appliances', 'C. Taking longer showers', 'D. Keeping windows open with AC on', 'B', 'Đoạn văn liệt kê "using energy-efficient appliances" là một cách tiết kiệm năng lượng.'),
    mcq(24, 'Reading: True/False', 'Q24: In Vietnam, electricity demand is highest during the winter months.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu nhu cầu điện cao nhất vào "hot summer months", không phải mùa đông.'),
    mcq(25, 'Reading: True/False', 'Q25: Installing solar panels can help lower electricity bills.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu lắp pin mặt trời "lowers their electricity bills".'),
    mcq(26, 'Reading: True/False', 'Q26: The passage states that saving energy is only about reducing costs.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu tiết kiệm năng lượng không chỉ về giảm chi phí mà còn để đảm bảo một hành tinh đáng sống.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "incandescent" in the passage most likely refers to:', 'A. A type of energy-saving bulb', 'B. A traditional type of light bulb', 'C. A solar panel component', 'D. A type of air conditioner', 'B', '"Incandescent bulb" là bóng đèn sợi đốt truyền thống, tiêu thụ nhiều điện hơn đèn LED.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of the passage?', 'A. Only the government should care about saving energy', 'B. Everyone can contribute to saving energy for a sustainable future', 'C. Solar panels are too expensive for most families', 'D. Vietnam does not need to worry about energy conservation', 'B', 'Bài đọc nhấn mạnh mọi người đều có thể góp phần tiết kiệm năng lượng vì một tương lai bền vững.'),
    textQ(29, 'Word Form', 'The government is encouraging the ____ of renewable energy sources across the country. (DEVELOP)', 'development', 'Cần danh từ sau mạo từ "the". "Develop" (động từ) -> "development" (danh từ).'),
    textQ(30, 'Word Form', 'Da Nang is a ____ city with a perfect blend of modern and traditional charm. (FASCINATE)', 'fascinating', 'Cần tính từ bổ nghĩa cho "city". "Fascinate" -> "fascinating" (hấp dẫn, lôi cuốn).'),
    textQ(31, 'Word Form', 'Using public transport instead of private cars helps to ____ air pollution in big cities. (REDUCTION)', 'reduce', 'Cần động từ trong cấu trúc "help to V". "Reduction" (danh từ) -> "reduce" (động từ).'),
    textQ(32, 'Word Form', 'The tourists were ____ impressed by the beauty of the Golden Bridge at sunrise. (DEEP)', 'deeply', 'Cần trạng từ bổ nghĩa cho tính từ "impressed". "Deep" -> "deeply".'),
    textQ(33, 'Word Form', 'Energy ____ is one of the most effective ways to protect the environment. (CONSERVE)', 'conservation', 'Cần danh từ làm chủ ngữ. "Conserve" (động từ) -> "conservation" (danh từ).'),
    textQ(34, 'Word Form', 'Many new hotels have been built to meet the ____ demand for accommodation in Da Nang. (INCREASE)', 'increasing', 'Cần tính từ bổ nghĩa cho "demand". "Increase" -> "increasing" (ngày càng tăng).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "generate /ˈdʒenəreɪt/ verb: to produce or create something, especially energy or electricity." Complete: The new wind farm can ____ enough electricity to power 10,000 homes.', 'generate', '"Generate electricity" (tạo ra điện) phù hợp với định nghĩa và ngữ cảnh trang trại gió.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "attract /əˈtrækt/ verb: to draw or pull someone or something towards oneself." Complete: The stunning beaches of Da Nang ____ millions of visitors every year.', 'attract', '"Attract visitors" (thu hút du khách) phù hợp với định nghĩa và ngữ cảnh.'),
    textQ(37, 'Sentence Transformation', 'They have installed solar panels on the roof of the building. (BEEN)\n→ Solar panels ____ on the roof of the building.', 'have been installed', 'Câu bị động thì hiện tại hoàn thành: S + have/has been + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'People say that Da Nang is the most livable city in Vietnam. (SAID)\n→ Da Nang ____ the most livable city in Vietnam.', 'is said to be', '"People say that S + V" → "S + is/are said + to V".'),
    textQ(39, 'Sentence Transformation', 'You should turn off electrical devices when not in use to save energy. (WERE)\n→ If I ____ off electrical devices when not in use to save energy.', 'were you, I would turn', '"You should V" → "If I were you, I would V" trong lời khuyên.'),
    textQ(40, 'Sentence Transformation', 'Despite the high cost, many families choose to install solar panels. (ALTHOUGH)\n→ ____ the cost is high, many families choose to install solar panels.', 'Although', '"Despite + N/V-ing" → "Although + clause".'),
  ];

  seedSingleTest_({
    test_id: '16',
    title: 'Đề thi thử số 16',
    description: 'Bám sát đề chính thức 2024 - Đà Nẵng và tiết kiệm năng lượng',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 17: Chủ đề Biến đổi khí hậu - Từ vựng môi trường và giải pháp xanh
// ======================================================================
function seedQuestionsTest17_() {
  var clozePassage = [
    "Climate change is one of the most pressing challenges (17)____ humanity today. Rising global temperatures",
    "have led to more frequent extreme weather (18)____, including floods, droughts, and heatwaves. Scientists warn",
    "that if greenhouse gas emissions are not reduced, the (19)____ will become increasingly severe. Many countries",
    "have committed to achieving net-zero carbon emissions (20)____ 2050 through green technologies and sustainable",
    "practices. Individuals can also contribute by reducing waste, using public transportation, and (21)____",
    "eco-friendly products. The fight against climate change requires collective (22)____ from governments,",
    "businesses, and individuals worldwide."
  ].join('\n\n');

  var readingPassage = [
    "Deforestation is one of the major drivers of climate change, responsible for about 10% of global greenhouse gas",
    "emissions. Every year, millions of hectares of forest are cleared for agriculture, logging, and urban development.",
    "Forests act as the planet's lungs, absorbing carbon dioxide and releasing oxygen. When trees are cut down, not",
    "only is this carbon-absorbing capacity lost, but the stored carbon is also released back into the atmosphere.",
    "In Vietnam, the government has launched reforestation programs aiming to plant one billion trees by 2025. Local",
    "communities are also involved in protecting mangroves, which serve as natural barriers against storms and rising",
    "sea levels. While progress has been made, experts stress that stronger enforcement of forest protection laws and",
    "greater public awareness are needed to reverse the trend of forest loss in the coming decades."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. reduced B. launched C. committed D. flooded', 'A. reduced /t/', 'B. launched /t/', 'C. committed /ɪd/', 'D. flooded /ɪd/', 'C', '"Committed" có âm cuối /t/ nên -ed đọc /ɪd/, trong khi "reduced" và "launched" đọc /t/ còn "flooded" cũng đọc /ɪd/. Khác biệt ở âm cuối.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. climate B. carbon C. certain D. conserve', 'A. climate /k/', 'B. carbon /k/', 'C. certain /s/', 'D. conserve /k/', 'C', '"Certain" có âm "c" đọc là /s/, các từ còn lại đọc là /k/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. climate B. absorb C. forest D. planet', 'A. climate', 'B. absorb', 'C. forest', 'D. planet', 'B', '"Absorb" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. emission B. government C. energy D. policy', 'A. emission', 'B. government', 'C. energy', 'D. policy', 'A', '"Emission" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The rise in ____ gases such as carbon dioxide and methane is the primary cause of global warming.', 'A. greenhouse', 'B. oxygen', 'C. natural', 'D. harmless', 'A', '"Greenhouse gases" (khí nhà kính) là thuật ngữ chuyên ngành chính xác.'),
    mcq(6, 'Grammar: Tenses', 'Scientists ____ the effects of climate change on marine ecosystems for over three decades.', 'A. studied', 'B. have been studying', 'C. study', 'D. were studying', 'B', 'Thì hiện tại hoàn thành tiếp diễn với "for over three decades" nhấn mạnh hành động liên tục.'),
    mcq(7, 'Phrasal Verbs', 'The environmental group is trying to ____ awareness about the importance of recycling.', 'A. bring about', 'B. put off', 'C. take away', 'D. give in', 'A', '"Bring about" = gây ra, tạo nên. "Bring about awareness" = nâng cao nhận thức.'),
    mcq(8, 'Prepositions', 'Many species are ____ risk of extinction due to the rapid loss of their natural habitats.', 'A. at', 'B. in', 'C. on', 'D. with', 'A', '"At risk of + N/V-ing" là cụm giới từ cố định.'),
    mcq(9, 'Grammar: Conditional', 'If the world ____ to burn fossil fuels at the current rate, temperatures will rise by 3°C by 2100.', 'A. continues', 'B. continued', 'C. will continue', 'D. would continue', 'A', 'Câu điều kiện loại 1: If + S + V (HTĐ), S + will + V.'),
    mcq(10, 'Grammar: Passive Voice', 'Millions of trees ____ in Vietnam each year as part of the national reforestation program.', 'A. plant', 'B. are planted', 'C. planted', 'D. have planted', 'B', 'Câu bị động thì hiện tại đơn: S + am/is/are + V3/ed. Chủ ngữ số nhiều nên dùng "are planted".'),
    mcq(11, 'Vocabulary: Word Choice', 'Switching to renewable energy is an important ____ toward a sustainable future.', 'A. step', 'B. stair', 'C. level', 'D. floor', 'A', '"An important step toward" = một bước quan trọng hướng tới. Là cụm từ cố định.'),
    mcq(12, 'Communication', 'Student A: "Do you think individual actions can help fight climate change?"\nStudent B: "____"', 'A. Absolutely. Every small action, like using less plastic, adds up.', 'B. I do not know what climate change is.', 'C. The weather is nice today.', 'D. Only scientists can solve this problem.', 'A', 'Câu trả lời khẳng định vai trò của hành động cá nhân trong cuộc chiến chống biến đổi khí hậu.'),
    mcq(13, 'Communication', 'Friend: "Why do not we start a recycling club at school?"\nYou: "____"', 'A. That sounds like a great idea! Let us talk to the teacher about it.', 'B. I am too busy to join any club.', 'C. Recycling is boring.', 'D. We do not have enough bins at school.', 'A', 'Phản hồi tích cực, ủng hộ ý tưởng thành lập câu lạc bộ tái chế.'),
    mcq(14, 'Public Signs', 'A sign showing three arrows forming a triangle (♻) means:', 'A. No parking', 'B. Recyclable material', 'C. Warning: toxic waste', 'D. Direction to follow', 'B', 'Biểu tượng ba mũi tên tam giác là ký hiệu quốc tế cho vật liệu có thể tái chế.'),
    mcq(15, 'Public Signs', 'What does a sign depicting a factory chimney with smoke inside a red crossed circle mean?', 'A. Factory ahead', 'B. Industrial zone', 'C. No emissions / pollution control zone', 'D. Smoking area', 'C', 'Biển báo ống khói có khói bị gạch chéo chỉ khu vực kiểm soát khí thải, cấm gây ô nhiễm.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'Many countries have promised ____ their carbon emissions by at least 30% by 2030.', 'A. cut', 'B. to cut', 'C. cutting', 'D. cutted', 'B', '"Promise + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "...one of the most pressing challenges (17)____ humanity today."', 'A. facing', 'B. faced', 'C. faces', 'D. face', 'A', '"Challenges facing humanity" = những thách thức mà nhân loại đang đối mặt. Rút gọn mệnh đề quan hệ.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...extreme weather (18)____, including floods, droughts, and heatwaves."', 'A. events', 'B. incidents', 'C. accidents', 'D. cases', 'A', '"Extreme weather events" = các hiện tượng thời tiết cực đoan, là cụm từ chuyên ngành.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...the (19)____ will become increasingly severe."', 'A. consequences', 'B. reasons', 'C. causes', 'D. sources', 'A', '"Consequences" (hậu quả) phù hợp với ngữ cảnh về tác động của biến đổi khí hậu.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...net-zero carbon emissions (20)____ 2050..."', 'A. by', 'B. until', 'C. in', 'D. on', 'A', '"By + mốc thời gian" = trước hoặc vào thời điểm đó. "By 2050" = trước năm 2050.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...using public transportation, and (21)____ eco-friendly products."', 'A. supporting', 'B. support', 'C. supported', 'D. supports', 'A', '"Supporting" song song với "reducing" và "using" (cấu trúc song song với V-ing).'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...requires collective (22)____ from governments, businesses, and individuals..."', 'A. action', 'B. act', 'C. acting', 'D. actor', 'A', '"Collective action" = hành động tập thể, là cụm danh từ cố định.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, deforestation is responsible for about what percentage of global greenhouse gas emissions?', 'A. About 5%', 'B. About 10%', 'C. About 25%', 'D. About 50%', 'B', 'Đoạn văn nêu: "responsible for about 10% of global greenhouse gas emissions".'),
    mcq(24, 'Reading: True/False', 'Q24: Forests absorb carbon dioxide and release oxygen, acting as the planet\'s lungs.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu rõ: "Forests act as the planet\'s lungs, absorbing carbon dioxide and releasing oxygen."'),
    mcq(25, 'Reading: True/False', 'Q25: When trees are cut down, their stored carbon remains safely in the soil forever.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu carbon được lưu trữ bị "released back into the atmosphere" khi cây bị chặt.'),
    mcq(26, 'Reading: True/False', 'Q26: Vietnam has a goal of planting one billion trees by 2025.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "planting one billion trees by 2025".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "mangroves" in the passage most likely refers to:', 'A. Types of tropical coastal trees that grow in saltwater', 'B. Mountain forests at high altitudes', 'C. Artificial forests in urban areas', 'D. Fruit orchards in the Mekong Delta', 'A', '"Mangroves" (rừng ngập mặn) là loại rừng cây nhiệt đới ven biển, có khả năng chắn bão và sóng.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of this passage?', 'A. To describe the beauty of Vietnamese forests', 'B. To explain the impact of deforestation and the need for forest protection', 'C. To promote tourism in Vietnamese national parks', 'D. To compare forests in different countries', 'B', 'Bài đọc mô tả tác động của phá rừng đối với biến đổi khí hậu và sự cần thiết của việc bảo vệ rừng.'),
    textQ(29, 'Word Form', 'The ____ of polar ice caps is one of the most visible effects of global warming. (MELT)', 'melting', 'Cần danh từ (danh động từ) sau mạo từ "The". "Melt" -> "melting".'),
    textQ(30, 'Word Form', 'Governments need to take ____ action to reduce the impact of climate change. (DECIDE)', 'decisive', 'Cần tính từ bổ nghĩa cho danh từ "action". "Decide" -> "decisive" (quyết đoán).'),
    textQ(31, 'Word Form', 'The ____ of many animal species is directly linked to climate change and habitat loss. (EXTINCT)', 'extinction', 'Cần danh từ sau mạo từ "The". "Extinct" (tính từ) -> "extinction" (danh từ).'),
    textQ(32, 'Word Form', 'Using renewable energy sources can ____ reduce our dependence on fossil fuels. (SIGNIFICANT)', 'significantly', 'Cần trạng từ bổ nghĩa cho động từ "reduce". "Significant" -> "significantly".'),
    textQ(33, 'Word Form', 'Environmental ____ has become a major concern for governments around the world. (DEGRADE)', 'degradation', 'Cần danh từ sau tính từ "Environmental". "Degrade" -> "degradation" (sự suy thoái).'),
    textQ(34, 'Word Form', 'The ____ of forests is essential for maintaining biodiversity and regulating the climate. (PRESERVE)', 'preservation', 'Cần danh từ sau mạo từ "The". "Preserve" -> "preservation" (sự bảo tồn).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "emit /ɪˈmɪt/ verb: to send out gas, heat, light, or sound into the environment." Complete: Factories that ____ large amounts of CO2 should be required to pay environmental taxes.', 'emit', '"Emit CO2" (thải khí CO2) phù hợp với định nghĩa và ngữ cảnh nhà máy thải khí thải.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "adapt /əˈdæpt/ verb: to change or adjust to new conditions or environments." Complete: Coastal communities must ____ to rising sea levels caused by global warming.', 'adapt', '"Adapt to" (thích nghi với) phù hợp với định nghĩa và ngữ cảnh biến đổi khí hậu.'),
    textQ(37, 'Sentence Transformation', 'People say that climate change is the biggest threat to our planet. (SAID)\n→ Climate change ____ the biggest threat to our planet.', 'is said to be', '"People say that S + V" → "S + is/are said + to V".'),
    textQ(38, 'Sentence Transformation', 'We must take action now, or the consequences will be disastrous. (IF)\n→ ____ now, the consequences will be disastrous.', 'If we do not take action', '"V + or + clause" → "If + S + do/does not V, + clause".'),
    textQ(39, 'Sentence Transformation', 'They have planted one million trees in the national park since last year. (BEEN)\n→ One million trees ____ in the national park since last year.', 'have been planted', 'Câu bị động thì hiện tại hoàn thành: S + have/has been + V3/ed.'),
    textQ(40, 'Sentence Transformation', '"Reduce your carbon footprint by walking or cycling more," the expert said. (ADVISED)\n→ The expert ____ carbon footprint by walking or cycling more.', 'advised us to reduce our', '"V + O," S said → S + advised + O + to V.'),
  ];

  seedSingleTest_({
    test_id: '17',
    title: 'Đề thi thử số 17',
    description: 'Chủ đề Biến đổi khí hậu - Từ vựng môi trường và giải pháp xanh',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 18: Chủ đề An toàn thực phẩm - Đọc hiểu về nông nghiệp hữu cơ
// ======================================================================
function seedQuestionsTest18_() {
  var clozePassage = [
    "Organic farming has gained (17)____ worldwide as consumers become more concerned about food safety and health.",
    "Unlike conventional farming, organic agriculture avoids the use (18)____ synthetic pesticides, chemical fertilizers,",
    "and genetically modified organisms. Organic farmers rely (19)____ natural methods such as crop rotation, composting,",
    "and biological pest control to maintain soil fertility. Although organic food is often (20)____ expensive than",
    "conventionally grown food, many people believe the health and environmental benefits justify the higher cost.",
    "In Vietnam, the demand (21)____ organic products is growing, particularly in urban areas. Experts recommend that",
    "consumers look for certified organic labels when shopping to (22)____ they are buying genuine organic products."
  ].join('\n\n');

  var readingPassage = [
    "Food safety has become an increasingly important concern for consumers around the world. Every year, millions",
    "of people become ill from eating contaminated food, according to the World Health Organization. Common causes",
    "of foodborne illnesses include bacteria such as Salmonella and E. coli, viruses, parasites, and chemical",
    "contaminants. In Vietnam, food safety has been a hot topic in the media, with frequent reports of unsafe food",
    "practices in markets and restaurants. The government has responded by strengthening food safety regulations and",
    "increasing inspections of food production facilities. However, experts say that consumers also have a role to play.",
    "Simple habits such as washing hands before preparing food, cooking meat thoroughly, checking expiration dates,",
    "and buying from trusted sources can significantly reduce the risk of food poisoning. Education about food safety",
    "should start at a young age so that children grow up understanding the importance of safe food practices."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -s/-es endings', 'Choose the word whose underlined part is pronounced differently: A. farms B. foods C. markets D. vegetables', 'A. farms /z/', 'B. foods /z/', 'C. markets /s/', 'D. vegetables /z/', 'C', '"Markets" có âm cuối /t/ vô thanh nên -s đọc là /s/, các từ còn lại đọc /z/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. organic B. contain C. vitamin D. pesticide', 'A. organic /ɔː/', 'B. contain /ə/', 'C. vitamin /aɪ/', 'D. pesticide /e/', 'C', '"Vitamin" có âm "i" đọc là /aɪ/ (Anh-Anh), các từ còn lại có âm khác ở âm tiết đầu.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. farmer B. healthy C. produce D. harvest', 'A. farmer', 'B. healthy', 'C. produce (n)', 'D. harvest', 'C', '"Produce" (nông sản) nhấn âm 2 (danh từ), các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. chemical B. organic C. pesticide D. nutrient', 'A. chemical', 'B. organic', 'C. pesticide', 'D. nutrient', 'B', '"Organic" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'Farmers should avoid using too many ____ because they can harm both human health and the environment.', 'A. nutrients', 'B. pesticides', 'C. vitamins', 'D. fertilizers (organic)', 'B', '"Pesticides" (thuốc trừ sâu hóa học) có thể gây hại cho sức khỏe và môi trường.'),
    mcq(6, 'Grammar: Tenses', 'In recent years, the demand for organic food ____ significantly in urban areas of Vietnam.', 'A. increases', 'B. increased', 'C. has increased', 'D. is increasing', 'C', 'Thì hiện tại hoàn thành với "In recent years".'),
    mcq(7, 'Phrasal Verbs', 'The food inspector decided to ____ the restaurant after finding unsanitary conditions in the kitchen.', 'A. close down', 'B. take up', 'C. give away', 'D. put on', 'A', '"Close down" = đóng cửa. Phù hợp ngữ cảnh thanh tra đóng cửa nhà hàng vì vi phạm vệ sinh.'),
    mcq(8, 'Prepositions', 'Many consumers are willing to pay more ____ organic food because they believe it is healthier.', 'A. for', 'B. on', 'C. with', 'D. at', 'A', '"Pay for + N" = trả tiền cho cái gì.'),
    mcq(9, 'Grammar: Comparison', 'Organic vegetables are typically ____ than conventionally grown ones, but many people prefer them.', 'A. more expensive', 'B. most expensive', 'C. expensive', 'D. as expensive', 'A', 'So sánh hơn với tính từ dài: more + adj + than.'),
    mcq(10, 'Grammar: Relative Clause', 'Consumers ____ buy organic food often do so because they are concerned about pesticides in their diet.', 'A. who', 'B. which', 'C. whom', 'D. where', 'A', '"Who" thay cho "consumers" (chỉ người) làm chủ ngữ trong mệnh đề quan hệ.'),
    mcq(11, 'Vocabulary: Word Choice', 'Always check the ____ date before buying packaged food to ensure it is still fresh.', 'A. expiration', 'B. production', 'C. distribution', 'D. preparation', 'A', '"Expiration date" = hạn sử dụng. Là thuật ngữ chuẩn trong an toàn thực phẩm.'),
    mcq(12, 'Communication', 'Customer: "Is this vegetable grown without chemicals?"\nShop assistant: "____"', 'A. Yes, it is 100% organic and certified.', 'B. I do not eat vegetables.', 'C. The price is on the tag.', 'D. The store closes at 9 p.m.', 'A', 'Câu trả lời xác nhận rau được trồng hữu cơ, phù hợp với câu hỏi của khách hàng.'),
    mcq(13, 'Communication', 'Mother: "Remember to wash your hands before eating."\nChild: "____"', 'A. OK, Mom. I will do it right now.', 'B. I am not hungry.', 'C. The food looks good.', 'D. Can I have dessert first?', 'A', 'Câu trả lời ngoan ngoãn, vâng lời mẹ về việc rửa tay trước khi ăn.'),
    mcq(14, 'Public Signs', 'A sign in a restaurant showing a person washing hands under a faucet means:', 'A. No entry', 'B. Employees must wash hands before returning to work', 'C. Fire exit', 'D. Drinking water available here', 'B', 'Biển báo rửa tay trong nhà hàng nhắc nhở nhân viên phải rửa tay trước khi làm việc.'),
    mcq(15, 'Public Signs', 'A sign showing a crossed-out cockroach in a food preparation area means:', 'A. Pest control — no pests allowed', 'B. Insect museum', 'C. Pet shop nearby', 'D. Outdoor dining area', 'A', 'Biển báo côn trùng bị gạch chéo chỉ khu vực kiểm soát dịch hại, không được có côn trùng.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The doctor recommended ____ more fresh fruits and vegetables for a healthier diet.', 'A. eat', 'B. to eat', 'C. eating', 'D. eaten', 'C', '"Recommend + V-ing" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Organic farming has gained (17)____ worldwide..."', 'A. popularity', 'B. popular', 'C. popularly', 'D. popularize', 'A', 'Cần danh từ sau động từ "gained". "Popular" -> "popularity" (sự phổ biến).'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...avoids the use (18)____ synthetic pesticides..."', 'A. of', 'B. for', 'C. with', 'D. by', 'A', '"The use of + N" = việc sử dụng cái gì.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "Organic farmers rely (19)____ natural methods..."', 'A. on', 'B. in', 'C. at', 'D. to', 'A', '"Rely on + N" = dựa vào, phụ thuộc vào.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...organic food is often (20)____ expensive than conventionally grown food..."', 'A. more', 'B. most', 'C. much', 'D. as', 'A', '"More expensive than" là cấu trúc so sánh hơn với tính từ dài.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...the demand (21)____ organic products is growing..."', 'A. for', 'B. of', 'C. in', 'D. with', 'A', '"Demand for + N" = nhu cầu về cái gì.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...look for certified organic labels when shopping to (22)____ they are buying genuine..."', 'A. ensure', 'B. assure', 'C. insure', 'D. reassure', 'A', '"Ensure (that) + clause" = đảm bảo rằng. Phù hợp nhất về nghĩa.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, what is a common cause of foodborne illnesses?', 'A. Exercise', 'B. Bacteria such as Salmonella and E. coli', 'C. Drinking too much water', 'D. Eating fresh vegetables', 'B', 'Đoạn văn nêu: "bacteria such as Salmonella and E. coli" là nguyên nhân phổ biến.'),
    mcq(24, 'Reading: True/False', 'Q24: According to the WHO, millions of people become ill from contaminated food every year.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "millions of people become ill from eating contaminated food, according to the WHO".'),
    mcq(25, 'Reading: True/False', 'Q25: In Vietnam, the government has not taken any action on food safety issues.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu chính phủ đã "strengthening food safety regulations and increasing inspections".'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, only the government is responsible for food safety.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nói người tiêu dùng cũng có vai trò trong việc đảm bảo an toàn thực phẩm qua các thói quen hàng ngày.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "contaminated" in the passage is closest in meaning to:', 'A. clean', 'B. polluted or infected', 'C. delicious', 'D. fresh', 'B', '"Contaminated" = bị ô nhiễm, nhiễm bẩn, nhiễm khuẩn. Đồng nghĩa với "polluted" hoặc "infected".'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of this passage?', 'A. Food safety is the government\'s only responsibility', 'B. Everyone has a role to play in ensuring food safety', 'C. Only restaurants need to worry about food safety', 'D. Food poisoning is not a serious problem', 'B', 'Bài đọc nhấn mạnh cả chính phủ và người tiêu dùng đều có vai trò trong việc đảm bảo an toàn thực phẩm.'),
    textQ(29, 'Word Form', 'The ____ of organic vegetables has increased significantly in big cities. (CONSUME)', 'consumption', 'Cần danh từ sau mạo từ "The". "Consume" (động từ) -> "consumption" (danh từ).'),
    textQ(30, 'Word Form', 'Always store raw meat ____ from cooked food to avoid cross-contamination. (SEPARATE)', 'separately', 'Cần trạng từ bổ nghĩa cho động từ "store". "Separate" -> "separately".'),
    textQ(31, 'Word Form', 'Many food ____ are now switching to organic methods to meet consumer demand. (PRODUCE)', 'producers', 'Cần danh từ chỉ người. "Produce" -> "producer" (nhà sản xuất), số nhiều thêm -s.'),
    textQ(32, 'Word Form', 'Proper food ____ is essential to prevent the growth of harmful bacteria. (STORE)', 'storage', 'Cần danh từ sau tính từ "Proper". "Store" (động từ) -> "storage" (danh từ).'),
    textQ(33, 'Word Form', 'The government has introduced ____ regulations to ensure food safety in all markets. (STRICT)', 'stricter', 'Cần tính từ so sánh hơn (có "more" ẩn hoặc so sánh ngắn). "Strict" -> "stricter" / "more strict".'),
    textQ(34, 'Word Form', 'Eating a ____ diet with plenty of fruits and vegetables helps boost your immune system. (BALANCE)', 'balanced', 'Cần tính từ bổ nghĩa cho "diet". "Balance" -> "balanced" (cân bằng).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "inspect /ɪnˈspekt/ verb: to examine something carefully and officially to ensure standards are met." Complete: Health officials regularly ____ restaurants to check their hygiene standards.', 'inspect', '"Inspect" (kiểm tra, thanh tra) phù hợp với định nghĩa và ngữ cảnh kiểm tra vệ sinh nhà hàng.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "preserve /prɪˈzɜːrv/ verb: to keep something in its original state or in good condition." Complete: Salt has been used for centuries to ____ food and prevent it from spoiling.', 'preserve', '"Preserve food" (bảo quản thực phẩm) phù hợp với định nghĩa và ngữ cảnh lịch sử dùng muối.'),
    textQ(37, 'Sentence Transformation', 'Although organic food is expensive, many families still choose to buy it. (HOWEVER)\n→ Organic food is expensive. ____, many families still choose to buy it.', 'However', '"Although + clause" → "Clause. However, + clause".'),
    textQ(38, 'Sentence Transformation', 'She started growing her own vegetables two years ago. (BEEN)\n→ She ____ her own vegetables for two years.', 'has been growing', '"S + started V-ing + time ago" → "S + have/has been V-ing + for + time".'),
    textQ(39, 'Sentence Transformation', '"Do not eat food that has passed its expiration date," the health expert warned. (WARNED)\n→ The health expert ____ food that had passed its expiration date.', 'warned us not to eat', 'Câu tường thuật với "warn": S + warned + O + (not) to V.'),
    textQ(40, 'Sentence Transformation', 'It is essential to wash fruits and vegetables before eating them. (MUST)\n→ Fruits and vegetables ____ before being eaten.', 'must be washed', '"It is essential to V" → "S + must be + V3/ed" (bị động).'),
  ];

  seedSingleTest_({
    test_id: '18',
    title: 'Đề thi thử số 18',
    description: 'Chủ đề An toàn thực phẩm - Đọc hiểu về nông nghiệp hữu cơ',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 19: Chủ đề Trí tuệ nhân tạo (AI) - Từ vựng công nghệ tương lai
// ======================================================================
function seedQuestionsTest19_() {
  var clozePassage = [
    "Artificial Intelligence (AI) is transforming nearly every (17)____ of modern life, from healthcare and education",
    "to transportation and entertainment. AI systems are designed to perform tasks (18)____ normally require human",
    "intelligence, such as recognizing speech, making decisions, and translating languages. In healthcare, AI helps",
    "doctors diagnose diseases more (19)____ and develop personalized treatment plans. However, the rapid development",
    "of AI also raises important (20)____ concerns about privacy, job displacement, and bias in algorithms. Experts",
    "emphasize the need for responsible AI development that prioritizes transparency and human well-being. (21)____",
    "many challenges remain, the potential benefits of AI are enormous. The key is to find ways (22)____ harness AI",
    "for good while minimizing its risks."
  ].join('\n\n');

  var readingPassage = [
    "ChatGPT and similar AI chatbots have taken the world by storm since their introduction. These large language",
    "models are trained on vast amounts of text data and can generate human-like responses to almost any question.",
    "In education, students have started using AI tools to help with homework, write essays, and even learn new",
    "languages. While some teachers worry that AI could encourage cheating, others see it as a powerful educational",
    "tool that can provide personalized tutoring to students who might not otherwise have access to extra help.",
    "In the workplace, AI is automating repetitive tasks, freeing up employees to focus on more creative and strategic",
    "work. However, there are legitimate concerns about AI replacing certain jobs entirely. Economists predict that",
    "while some roles will disappear, new jobs we cannot yet imagine will also be created. The challenge for society",
    "is to ensure that workers have the skills and support they need to adapt to this rapidly changing landscape."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. artificial B. data C. machine D. algorithm', 'A. artificial /ɑːr/', 'B. data /eɪ/', 'C. machine /ə/', 'D. algorithm /æ/', 'C', '"Machine" có âm "a" đọc là /ʃ/, khác với các từ còn lại. Xét âm tiết được gạch chân.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. technology B. chatbot C. machine D. challenge', 'A. technology /k/', 'B. chatbot /tʃ/', 'C. machine /ʃ/', 'D. challenge /tʃ/', 'C', '"Machine" có âm "ch" đọc là /ʃ/, các từ còn lại có âm /k/ hoặc /tʃ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. robot B. machine C. data D. software', 'A. robot', 'B. machine', 'C. data', 'D. software', 'B', '"Machine" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. algorithm B. intelligence C. automate D. analyze', 'A. algorithm', 'B. intelligence', 'C. automate', 'D. analyze', 'B', '"Intelligence" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The new AI software can ____ large amounts of data in just a few seconds, making it invaluable for research.', 'A. process', 'B. produce', 'C. prevent', 'D. pretend', 'A', '"Process data" (xử lý dữ liệu) là cụm từ chuyên ngành công nghệ thông tin.'),
    mcq(6, 'Grammar: Tenses', 'AI technology ____ rapidly over the past decade, changing the way we live and work.', 'A. evolved', 'B. has evolved', 'C. evolves', 'D. was evolving', 'B', 'Thì hiện tại hoàn thành với "over the past decade".'),
    mcq(7, 'Phrasal Verbs', 'The IT team is working to ____ a solution to the software bug before the system launches.', 'A. come up with', 'B. put up with', 'C. look down on', 'D. make up for', 'A', '"Come up with" = nghĩ ra, tìm ra. "Come up with a solution" = tìm ra giải pháp.'),
    mcq(8, 'Prepositions', 'Many people are worried ____ the impact of AI on job security and employment opportunities.', 'A. about', 'B. for', 'C. with', 'D. at', 'A', '"Worried about + N/V-ing" = lo lắng về điều gì.'),
    mcq(9, 'Grammar: Conditional', 'If AI systems ____ properly regulated, they could cause serious harm to society.', 'A. are not', 'B. were not', 'C. will not be', 'D. would not be', 'A', 'Câu điều kiện loại 1: If + S + V (HTĐ), S + could/will + V.'),
    mcq(10, 'Grammar: Passive Voice', 'Many routine tasks ____ by AI in the near future, according to technology experts.', 'A. will automate', 'B. will be automated', 'C. automate', 'D. are automating', 'B', 'Câu bị động thì tương lai đơn: S + will be + V3/ed.'),
    mcq(11, 'Vocabulary: Word Choice', 'Virtual assistants like Siri and Alexa use AI ____ to understand and respond to voice commands.', 'A. algorithms', 'B. equations', 'C. formulas', 'D. codes', 'A', '"AI algorithms" (thuật toán AI) là thuật ngữ chính xác trong lĩnh vực trí tuệ nhân tạo.'),
    mcq(12, 'Communication', 'Student A: "Do you think AI will replace teachers in the future?"\nStudent B: "____"', 'A. I think AI can assist teachers, but it cannot replace the human connection in education.', 'B. I do not like computers.', 'C. My teacher is very strict.', 'D. AI stands for Artificial Intelligence.', 'A', 'Câu trả lời có lập luận cân nhắc về vai trò của AI trong giáo dục, phù hợp với câu hỏi.'),
    mcq(13, 'Communication', 'Colleague: "I am worried that AI might take over my job soon."\nYou: "____"', 'A. Instead of worrying, why not learn some AI-related skills to stay competitive?', 'B. You are right. You should just quit now.', 'C. AI is not important in our industry.', 'D. I do not care about technology at all.', 'A', 'Lời khuyên tích cực, khuyến khích học kỹ năng mới để thích nghi thay vì chỉ lo lắng.'),
    mcq(14, 'Public Signs', 'A sign at a data center showing a server rack with a lock symbol means:', 'A. Free Wi-Fi zone', 'B. Authorized personnel only — restricted access', 'C. Public computer lab', 'D. Internet café', 'B', 'Biển báo máy chủ có ổ khóa chỉ khu vực hạn chế, chỉ nhân viên được phép vào.'),
    mcq(15, 'Public Signs', 'What does a sign showing a camera with "CCTV in Operation" mean?', 'A. No photography', 'B. Video surveillance zone', 'C. Photo studio', 'D. Camera repair shop', 'B', 'Biển báo "CCTV in Operation" cho biết khu vực đang được giám sát bằng camera an ninh.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'Many companies are considering ____ AI to improve their customer service systems.', 'A. adopt', 'B. to adopt', 'C. adopting', 'D. adopted', 'C', '"Consider + V-ing" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "AI is transforming nearly every (17)____ of modern life..."', 'A. aspect', 'B. respect', 'C. inspect', 'D. suspect', 'A', '"Every aspect of" = mọi khía cạnh của. Phù hợp nhất về nghĩa và cách dùng.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...perform tasks (18)____ normally require human intelligence..."', 'A. that', 'B. who', 'C. whom', 'D. where', 'A', '"That" thay cho "tasks" (chỉ vật) làm chủ ngữ trong mệnh đề quan hệ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...helps doctors diagnose diseases more (19)____..."', 'A. accurate', 'B. accurately', 'C. accuracy', 'D. more accurate', 'B', 'Cần trạng từ bổ nghĩa cho động từ "diagnose". "Accurate" -> "accurately".'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...raises important (20)____ concerns about privacy..."', 'A. ethical', 'B. ethnic', 'C. external', 'D. eternal', 'A', '"Ethical concerns" = những lo ngại về đạo đức. Phù hợp trong ngữ cảnh AI.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "(21)____ many challenges remain, the potential benefits of AI are enormous."', 'A. Although', 'B. Because', 'C. If', 'D. Unless', 'A', '"Although" diễn tả sự tương phản: mặc dù còn nhiều thách thức nhưng lợi ích rất lớn.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...find ways (22)____ harness AI for good while minimizing its risks."', 'A. to', 'B. of', 'C. for', 'D. in', 'A', '"Ways to V" = những cách để làm gì. Cấu trúc cố định.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, how do some teachers view AI tools in education?', 'A. As a threat that should be banned', 'B. As a powerful tool that can provide personalized tutoring', 'C. As a replacement for all teachers', 'D. As something that has no use in classrooms', 'B', 'Đoạn văn nêu một số giáo viên xem AI là "a powerful educational tool that can provide personalized tutoring".'),
    mcq(24, 'Reading: True/False', 'Q24: AI chatbots like ChatGPT are trained on small amounts of text data.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu chúng được huấn luyện trên "vast amounts of text data", không phải lượng nhỏ.'),
    mcq(25, 'Reading: True/False', 'Q25: According to economists, AI will only destroy jobs without creating any new ones.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu một số công việc sẽ biến mất nhưng "new jobs we cannot yet imagine will also be created".'),
    mcq(26, 'Reading: True/False', 'Q26: The passage suggests that workers need support to adapt to the changing job landscape.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu thách thức là đảm bảo "workers have the skills and support they need to adapt".'),
    mcq(27, 'Reading: MCQ', 'Q27: The phrase "taken the world by storm" in the passage most likely means:', 'A. Caused a natural disaster', 'B. Become suddenly very popular and successful', 'C. Created a lot of problems everywhere', 'D. Been banned in many countries', 'B', '"Taken the world by storm" = gây bão toàn cầu, trở nên cực kỳ phổ biến một cách nhanh chóng.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the author\'s overall perspective on AI in this passage?', 'A. AI should be completely banned', 'B. AI brings both opportunities and challenges that society must manage carefully', 'C. AI will solve all of humanity\'s problems', 'D. AI is not important enough to discuss', 'B', 'Tác giả thể hiện quan điểm cân bằng: AI mang lại cơ hội và thách thức, xã hội cần quản lý cẩn thận.'),
    textQ(29, 'Word Form', 'The ____ of artificial intelligence has brought both opportunities and challenges to society. (DEVELOP)', 'development', 'Cần danh từ sau mạo từ "The". "Develop" (động từ) -> "development" (danh từ).'),
    textQ(30, 'Word Form', 'AI can help doctors diagnose diseases ____ by analyzing medical images in seconds. (ACCURATE)', 'accurately', 'Cần trạng từ bổ nghĩa cho động từ "diagnose". "Accurate" -> "accurately".'),
    textQ(31, 'Word Form', 'The ____ of AI in education has raised questions about academic honesty. (APPLY)', 'application', 'Cần danh từ sau mạo từ "The". "Apply" (động từ) -> "application" (danh từ).'),
    textQ(32, 'Word Form', 'We need to be ____ about the information we receive from AI chatbots. (CRITIC)', 'critical', 'Cần tính từ sau "be". "Critic" (danh từ) -> "critical" (tính từ).'),
    textQ(33, 'Word Form', 'AI is a ____ tool that will continue to shape the future of humanity. (POWER)', 'powerful', 'Cần tính từ bổ nghĩa cho "tool". "Power" (danh từ) -> "powerful" (tính từ).'),
    textQ(34, 'Word Form', 'The company is investing in AI ____ to stay ahead of its competitors. (INNOVATE)', 'innovation', 'Cần danh từ sau "AI". "Innovate" (động từ) -> "innovation" (danh từ).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "simulate /ˈsɪmjʊleɪt/ verb: to create a representation or model of a system or process." Complete: Computer programs can ____ complex weather patterns to help predict future climate conditions.', 'simulate', '"Simulate" (mô phỏng) phù hợp với định nghĩa và ngữ cảnh mô phỏng thời tiết.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "compute /kəmˈpjuːt/ verb: to calculate or process data using a computer." Complete: Modern supercomputers can ____ billions of operations per second.', 'compute', '"Compute operations" (tính toán các phép tính) phù hợp với định nghĩa và ngữ cảnh siêu máy tính.'),
    textQ(37, 'Sentence Transformation', 'Scientists have developed a new AI model that can translate languages instantly. (BEEN)\n→ A new AI model ____ that can translate languages instantly.', 'has been developed by scientists', 'Câu bị động thì hiện tại hoàn thành: S + have/has been + V3/ed + by O.'),
    textQ(38, 'Sentence Transformation', '"We are launching the new AI feature next month," the CEO announced. (ANNOUNCED)\n→ The CEO ____ the new AI feature the following month.', 'announced that they were launching', 'Câu tường thuật: lùi thì "are launching" → "were launching", "next month" → "the following month".'),
    textQ(39, 'Sentence Transformation', 'Learning about AI is important, but developing soft skills is equally necessary. (JUST)\n→ Developing soft skills is ____ learning about AI.', 'just as important as', '"Equally necessary" → "just as + adj + as". Cấu trúc so sánh ngang bằng.'),
    textQ(40, 'Sentence Transformation', 'If AI is not carefully regulated, it could cause serious problems. (UNLESS)\n→ ____ carefully regulated, AI could cause serious problems.', 'Unless it is', '"If...not" → "Unless". Unless + S + V = If + S + do/does not + V.'),
  ];

  seedSingleTest_({
    test_id: '19',
    title: 'Đề thi thử số 19',
    description: 'Chủ đề Trí tuệ nhân tạo (AI) - Từ vựng công nghệ tương lai',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 20: Chủ đề Mạng xã hội & Truyền thông - Giao tiếp và ứng xử online
// ======================================================================
function seedQuestionsTest20_() {
  var clozePassage = [
    "Social media has fundamentally changed the (17)____ people communicate and share information in the 21st century.",
    "Platforms like Facebook, Instagram, and TikTok connect billions of users worldwide, (18)____ them to share moments",
    "and ideas instantly. While social media offers many benefits, (19)____ staying in touch with friends and accessing",
    "news, it also has significant drawbacks. Excessive use of social media can (20)____ to anxiety, depression, and",
    "reduced face-to-face social interactions. Experts recommend setting time limits and being mindful (21)____ the",
    "content we consume online. Parents should also guide their children on (22)____ to use social media responsibly."
  ].join('\n\n');

  var readingPassage = [
    "In the age of social media, the line between public and private life has become increasingly blurred. Every day,",
    "millions of people post photos, share opinions, and update their status on various platforms without fully",
    "considering the potential consequences. Once something is posted online, it can be almost impossible to completely",
    "remove it. Employers now routinely check applicants' social media profiles before making hiring decisions.",
    "Similarly, university admissions officers sometimes review candidates' online presence. On the positive side,",
    "social media has given ordinary people a platform to share their voices, build communities around shared interests,",
    "and even launch successful businesses. Influencers and content creators can reach global audiences with just a",
    "smartphone and an internet connection. The key to navigating the digital world is to think before you post and",
    "to use privacy settings wisely. Digital literacy — understanding how to use technology safely and responsibly —",
    "has become just as important as traditional literacy in today's connected world."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. posted B. shared C. liked D. followed', 'A. posted /ɪd/', 'B. shared /d/', 'C. liked /t/', 'D. followed /d/', 'A', '"Posted" có âm cuối /t/ nên -ed đọc /ɪd/, các từ còn lại đọc /d/ hoặc /t/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. media B. message C. content D. internet', 'A. media /iː/', 'B. message /e/', 'C. content /ɒ/', 'D. internet /ɪ/', 'A', '"Media" có âm "e" đọc là /iː/, các từ còn lại có âm /e/, /ɒ/, /ɪ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. online B. social C. TikTok D. Facebook', 'A. online', 'B. social', 'C. TikTok', 'D. Facebook', 'A', '"Online" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. Internet B. privacy C. computer D. digital', 'A. Internet', 'B. privacy', 'C. computer', 'D. digital', 'C', '"Computer" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'It is important to ____ information before sharing it on social media to avoid spreading fake news.', 'A. verify', 'B. ignore', 'C. delete', 'D. copy', 'A', '"Verify information" (xác minh thông tin) là hành động quan trọng trước khi chia sẻ trên mạng xã hội.'),
    mcq(6, 'Grammar: Tenses', 'In the past few years, TikTok ____ one of the most popular social media platforms among teenagers.', 'A. becomes', 'B. became', 'C. has become', 'D. is becoming', 'C', 'Thì hiện tại hoàn thành với "In the past few years".'),
    mcq(7, 'Phrasal Verbs', 'Many teenagers find it hard to ____ from social media, even during important exams.', 'A. stay away', 'B. take over', 'C. put down', 'D. give away', 'A', '"Stay away from + N" = tránh xa, không sử dụng. Phù hợp ngữ cảnh cai mạng xã hội.'),
    mcq(8, 'Prepositions', 'Spending too much time ____ social media can negatively affect your mental health and sleep quality.', 'A. on', 'B. in', 'C. at', 'D. for', 'A', '"Spend time on + N" = dành thời gian vào việc gì.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish my parents ____ more understanding about my use of social media for school projects.', 'A. are', 'B. were', 'C. will be', 'D. have been', 'B', 'Câu ước ở hiện tại: S + wish + S + V (quá khứ đơn). Dùng "were" cho mọi ngôi.'),
    mcq(10, 'Grammar: Relative Clause', 'The influencer ____ videos went viral last week has gained over one million new followers.', 'A. who', 'B. which', 'C. whose', 'D. whom', 'C', '"Whose" là đại từ quan hệ sở hữu, thay cho "the influencer\'s".'),
    mcq(11, 'Vocabulary: Word Choice', 'Cyberbullying is a serious ____ that affects millions of young people on social media platforms every day.', 'A. issue', 'B. solution', 'C. benefit', 'D. advantage', 'A', '"A serious issue" = một vấn đề nghiêm trọng. Phù hợp với ngữ cảnh bắt nạt trên mạng.'),
    mcq(12, 'Communication', 'Friend: "Hey, I saw some mean comments on your last post. Are you okay?"\nYou: "____"', 'A. Thanks for asking. I decided to block them and report the comments.', 'B. I do not use social media anymore.', 'C. What is your username?', 'D. Can you help me post a new photo?', 'A', 'Phản hồi phù hợp khi bạn bè quan tâm về bình luận tiêu cực trên mạng xã hội.'),
    mcq(13, 'Communication', 'Parent: "You spend too much time on your phone. I am worried about your grades."\nTeenager: "____"', 'A. I understand your concern, Mom. How about I set a daily screen time limit?', 'B. Leave me alone! You do not understand anything!', 'C. All my friends use phones even more than I do.', 'D. I do not care about my grades anyway.', 'A', 'Phản hồi trưởng thành, tôn trọng và đề xuất giải pháp khi bị phụ huynh nhắc nhở.'),
    mcq(14, 'Public Signs', 'What does the "@" symbol on a sign at a café usually mean?', 'A. Menu prices in dollars', 'B. Free Wi-Fi available', 'C. No phones allowed', 'D. Restroom location', 'B', 'Biểu tượng @ trên biển quán café thường thông báo có Wi-Fi miễn phí.'),
    mcq(15, 'Public Signs', 'A sign showing a smartphone with a red cross line through it in a cinema means:', 'A. Phone charging station available', 'B. Please turn off mobile phones', 'C. Free phone rental', 'D. Emergency phone location', 'B', 'Biển báo điện thoại bị gạch chéo trong rạp chiếu phim yêu cầu tắt điện thoại di động.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She refused ____ her password with anyone, even her closest friends.', 'A. share', 'B. to share', 'C. sharing', 'D. shared', 'B', '"Refuse + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Social media has fundamentally changed the (17)____ people communicate..."', 'A. way', 'B. method', 'C. style', 'D. path', 'A', '"The way + S + V" = cách mà ai đó làm gì. Cấu trúc cố định và tự nhiên nhất.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...connect billions of users worldwide, (18)____ them to share moments..."', 'A. allowing', 'B. allow', 'C. allowed', 'D. allows', 'A', '"Allowing" là phân từ hiện tại, rút gọn mệnh đề quan hệ chủ động.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...many benefits, (19)____ staying in touch with friends..."', 'A. such as', 'B. so that', 'C. because of', 'D. in order to', 'A', '"Such as" = ví dụ như, dùng để liệt kê các lợi ích.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "Excessive use of social media can (20)____ to anxiety..."', 'A. lead', 'B. cause', 'C. result', 'D. bring', 'A', '"Lead to + N" = dẫn đến điều gì. Cấu trúc cố định.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...being mindful (21)____ the content we consume online."', 'A. about', 'B. of', 'C. with', 'D. for', 'B', '"Be mindful of + N" = chú ý, để tâm đến điều gì.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...guide their children on (22)____ to use social media responsibly."', 'A. how', 'B. what', 'C. which', 'D. where', 'A', '"On how to V" = về cách để làm gì.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, why do employers check social media profiles of job applicants?', 'A. To find new friends', 'B. To help make hiring decisions', 'C. To sell them products', 'D. To invite them to events', 'B', 'Đoạn văn nêu: "Employers now routinely check applicants\' social media profiles before making hiring decisions."'),
    mcq(24, 'Reading: True/False', 'Q24: Once something is posted online, it is easy to completely remove it.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu: "it can be almost impossible to completely remove it."'),
    mcq(25, 'Reading: True/False', 'Q25: According to the passage, only negative aspects of social media are discussed.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn cũng đề cập đến mặt tích cực: "social media has given ordinary people a platform to share their voices".'),
    mcq(26, 'Reading: True/False', 'Q26: Digital literacy is considered as important as traditional literacy in today\'s world.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "Digital literacy...has become just as important as traditional literacy."'),
    mcq(27, 'Reading: MCQ', 'Q27: The phrase "think before you post" in the passage is best understood as:', 'A. Wait for someone else to post first', 'B. Consider the consequences before sharing content online', 'C. Only post when you are happy', 'D. Ask your parents before using social media', 'B', '"Think before you post" = suy nghĩ trước khi đăng, cân nhắc hậu quả trước khi chia sẻ nội dung.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of this passage?', 'A. To encourage people to delete all social media accounts', 'B. To highlight both the risks and benefits of social media and the importance of responsible use', 'C. To promote a specific social media platform', 'D. To argue that social media has no positive effects', 'B', 'Bài đọc cân bằng giữa rủi ro và lợi ích của mạng xã hội, nhấn mạnh tầm quan trọng của việc sử dụng có trách nhiệm.'),
    textQ(29, 'Word Form', 'Social media can be ____ if not used responsibly and in moderation. (HARM)', 'harmful', 'Cần tính từ sau "be". "Harm" (danh từ/động từ) -> "harmful" (tính từ).'),
    textQ(30, 'Word Form', 'Parents should ____ their children\'s online activities to protect them from potential dangers. (SUPERVISION)', 'supervise', 'Cần động từ sau "should". "Supervision" (danh từ) -> "supervise" (động từ).'),
    textQ(31, 'Word Form', 'Many young people are not fully ____ of the long-term consequences of oversharing online. (AWARENESS)', 'aware', 'Cần tính từ trong cấu trúc "be ____ of". "Awareness" (danh từ) -> "aware" (tính từ).'),
    textQ(32, 'Word Form', 'The ____ of news on social media can sometimes be faster than traditional media outlets. (SPREAD)', 'spread', 'Cần danh từ sau mạo từ "The". "Spread" có thể là danh từ chỉ sự lan truyền. Giữ nguyên.'),
    textQ(33, 'Word Form', 'He was ____ by the negative comments on his post and decided to take a break from social media. (AFFECT)', 'affected', 'Cần quá khứ phân từ trong câu bị động "was ____ by". "Affect" -> "affected" (bị ảnh hưởng).'),
    textQ(34, 'Word Form', 'Instagram is one of the most ____ social media platforms for sharing photos and short videos. (POPULARITY)', 'popular', 'Cần tính từ trong cấu trúc "one of the most + adj". "Popularity" -> "popular".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "block /blɒk/ verb: to stop someone from contacting you on a social media platform." Complete: If someone is harassing you online, the best first step is to ____ them immediately.', 'block', '"Block" (chặn) phù hợp với định nghĩa và ngữ cảnh ngăn chặn quấy rối trực tuyến.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "connect /kəˈnekt/ verb: to join or link people, things, or ideas together." Complete: Social media helps people ____ with friends and family members who live far away.', 'connect', '"Connect with" (kết nối với) phù hợp với định nghĩa và ngữ cảnh giữ liên lạc với người thân ở xa.'),
    textQ(37, 'Sentence Transformation', 'People should limit their time on social media to protect their mental health. (IF)\n→ People can protect their mental health ____ their time on social media.', 'if they limit', '"S + should V + to V" → "S + can V + if + S + V (HTĐ)".'),
    textQ(38, 'Sentence Transformation', 'She last checked her Facebook account three days ago. (HAS NOT)\n→ She ____ her Facebook account for three days.', 'has not checked', '"S + last V (QKĐ) + time ago" → "S + have/has not V3/ed + for + time".'),
    textQ(39, 'Sentence Transformation', '"Do not accept friend requests from strangers," the teacher told the students. (WARNED)\n→ The teacher ____ friend requests from strangers.', 'warned the students not to accept', 'Câu tường thuật với "warn": S + warned + O + not to V.'),
    textQ(40, 'Sentence Transformation', 'Using social media too much can cause mental health problems. (LEAD)\n→ Using social media too much can ____.', 'lead to mental health problems', '"Cause" → "lead to". "Lead to + N" = gây ra, dẫn đến.'),
  ];

  seedSingleTest_({
    test_id: '20',
    title: 'Đề thi thử số 20',
    description: 'Chủ đề Mạng xã hội & Truyền thông - Giao tiếp và ứng xử online',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 21: Chủ đề Lịch sử & Di tích - Đọc hiểu về di sản văn hóa Việt Nam
// ======================================================================
function seedQuestionsTest21_() {
  var clozePassage = [
    "Vietnam is home to eight UNESCO World Heritage Sites, reflecting the country's (17)____ history and cultural",
    "diversity. Ha Long Bay, recognized in 1994, is famous (18)____ its emerald waters and thousands of towering",
    "limestone islands. The Complex of Hue Monuments preserves the imperial legacy of the Nguyen Dynasty, the last",
    "ruling family of Vietnam. Hoi An Ancient Town, once a bustling trading (19)____, showcases a unique blend of",
    "Vietnamese, Chinese, and Japanese architectural influences. These heritage sites not (20)____ attract millions",
    "of tourists but also serve as important (21)____ of Vietnam's cultural identity. In recent years, the government",
    "has invested more in (22)____ and restoring these sites for future generations."
  ].join('\n\n');

  var readingPassage = [
    "The Temple of Literature in Hanoi is one of Vietnam's most iconic historical sites. Built in 1070 under the reign",
    "of Emperor Ly Thanh Tong, it was originally dedicated to Confucius and served as the country's first national",
    "university. For centuries, the brightest scholars from across the country came here to study literature, philosophy,",
    "and politics. The temple complex features beautiful traditional Vietnamese architecture, with peaceful courtyards,",
    "ancient trees, and the famous Stelae of Doctors — stone slabs mounted on turtle statues that list the names of",
    "outstanding graduates. Today, the Temple of Literature is a popular tourist attraction and a symbol of Vietnam's",
    "long-standing commitment to education. Before important exams, many students visit the temple to pray for good",
    "luck and touch the stone turtle heads, hoping to achieve academic success. The image of the Temple of Literature",
    "even appears on the back of the 100,000 Vietnamese dong banknote, underscoring its national significance."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. recognized B. dedicated C. constructed D. preserved', 'A. recognized /d/', 'B. dedicated /ɪd/', 'C. constructed /ɪd/', 'D. preserved /d/', 'B', '"Dedicated" có âm cuối /t/ nên -ed đọc /ɪd/, trong khi "recognized" và "preserved" đọc /d/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. ancient B. cultural C. century D. complex', 'A. ancient /eɪ/', 'B. cultural /ʌ/', 'C. century /e/', 'D. complex /ɒ/', 'B', '"Cultural" có âm "u" đọc là /ʌ/, các từ còn lại có âm khác. Xét âm tiết đầu.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. temple B. complex C. preserve D. palace', 'A. temple', 'B. complex', 'C. preserve', 'D. palace', 'C', '"Preserve" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. heritage B. imperial C. monument D. dynasty', 'A. heritage', 'B. imperial', 'C. monument', 'D. dynasty', 'B', '"Imperial" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The ancient citadel was ____ in the 19th century and still stands as a symbol of national pride.', 'A. constructed', 'B. destroyed', 'C. neglected', 'D. forgotten', 'A', '"Constructed" (được xây dựng) phù hợp với ngữ cảnh mô tả công trình cổ còn tồn tại.'),
    mcq(6, 'Grammar: Tenses', 'Hoi An Ancient Town ____ as a UNESCO World Heritage Site since 1999.', 'A. recognized', 'B. was recognized', 'C. has been recognized', 'D. recognizes', 'C', 'Thì hiện tại hoàn thành bị động với "since" + mốc thời gian.'),
    mcq(7, 'Phrasal Verbs', 'The local community is working hard to ____ the historic temple after years of neglect.', 'A. restore', 'B. break down', 'C. put off', 'D. give away', 'A', '"Restore" (khôi phục, trùng tu) phù hợp với ngữ cảnh phục dựng đền chùa cổ.'),
    mcq(8, 'Prepositions', 'The Imperial City in Hue was built ____ the reign of Emperor Gia Long in the early 19th century.', 'A. during', 'B. while', 'C. among', 'D. between', 'A', '"During the reign of" = dưới triều đại của. Là cụm giới từ lịch sử.'),
    mcq(9, 'Grammar: Passive Voice', 'The Temple of Literature ____ by thousands of tourists and students every year.', 'A. visits', 'B. is visited', 'C. visited', 'D. has visited', 'B', 'Câu bị động thì hiện tại đơn: S (số ít) + is + V3/ed.'),
    mcq(10, 'Grammar: Relative Clause', 'The archaeologists ____ discovered the ancient ruins were awarded medals by the government.', 'A. who', 'B. which', 'C. where', 'D. when', 'A', '"Who" thay cho "the archaeologists" (chỉ người) làm chủ ngữ mệnh đề quan hệ.'),
    mcq(11, 'Vocabulary: Word Choice', 'The Stelae of Doctors at the Temple of Literature ____ the names of scholars who passed the royal exams.', 'A. record', 'B. hide', 'C. erase', 'D. forget', 'A', '"Record" (ghi lại) phù hợp với chức năng của bia Tiến sĩ — ghi tên những người đỗ đạt.'),
    mcq(12, 'Communication', 'Tourist: "How old is this temple?"\nTour guide: "____"', 'A. It was built in 1070, so it is nearly a thousand years old.', 'B. I do not know the answer to that.', 'C. The entrance fee is 30,000 VND.', 'D. We are now at the main gate.', 'A', 'Hướng dẫn viên trả lời chính xác và thú vị về tuổi của ngôi đền.'),
    mcq(13, 'Communication', 'Visitor: "Is photography allowed inside the temple?"\nGuard: "____"', 'A. Yes, but please do not use flash as it can damage the artifacts.', 'B. No, this is a restaurant.', 'C. I am not a photographer.', 'D. The temple closes at 5 p.m.', 'A', 'Câu trả lời cho phép chụp ảnh nhưng kèm theo lưu ý bảo vệ di vật, phù hợp trong bảo tàng/di tích.'),
    mcq(14, 'Public Signs', 'A sign at a heritage site showing a camera with a flash symbol crossed out means:', 'A. Photography is encouraged', 'B. No flash photography', 'C. Video recording only', 'D. Camera rental available', 'B', 'Biển báo máy ảnh có tia sét bị gạch chéo nghĩa là cấm chụp ảnh có đèn flash.'),
    mcq(15, 'Public Signs', 'What does a sign showing a hand touching an object with a red cross mean at a museum?', 'A. Please touch the exhibits', 'B. Do not touch the exhibits', 'C. Interactive display', 'D. Hand sanitizer station', 'B', 'Biển báo bàn tay chạm vào đồ vật bị gạch chéo nghĩa là không được chạm vào hiện vật trưng bày.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The government plans ____ more funding into preserving historical sites across the country.', 'A. invest', 'B. to invest', 'C. investing', 'D. invested', 'B', '"Plan + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "...reflecting the country\'s (17)____ history and cultural diversity."', 'A. rich', 'B. poor', 'C. short', 'D. boring', 'A', '"Rich history" = lịch sử phong phú, giàu có. Phù hợp với ngữ cảnh UNESCO công nhận.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "Ha Long Bay...is famous (18)____ its emerald waters..."', 'A. for', 'B. about', 'C. with', 'D. at', 'A', '"Famous for + N" = nổi tiếng về điều gì.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...once a bustling trading (19)____..."', 'A. port', 'B. airport', 'C. station', 'D. stop', 'A', '"Trading port" = cảng thương mại. Hội An từng là thương cảng sầm uất.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "These heritage sites not (20)____ attract millions of tourists..."', 'A. only', 'B. just', 'C. also', 'D. even', 'A', 'Cấu trúc "not only...but also" = không những...mà còn.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...serve as important (21)____ of Vietnam\'s cultural identity."', 'A. reminders', 'B. remembers', 'C. memories', 'D. memorials', 'A', '"Reminders of" = những lời nhắc nhở về. Phù hợp nhất về nghĩa.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...the government has invested more in (22)____ and restoring these sites..."', 'A. preserving', 'B. destroying', 'C. ignoring', 'D. removing', 'A', '"Preserving and restoring" = bảo tồn và trùng tu. Song song về cấu trúc và phù hợp về nghĩa.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: When was the Temple of Literature built?', 'A. In 1010', 'B. In 1070', 'C. In 1107', 'D. In 1700', 'B', 'Đoạn văn nêu: "Built in 1070 under the reign of Emperor Ly Thanh Tong".'),
    mcq(24, 'Reading: True/False', 'Q24: The Temple of Literature was Vietnam\'s first national university.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "it served as the country\'s first national university."'),
    mcq(25, 'Reading: True/False', 'Q25: The Stelae of Doctors are made of wood and list the names of royal family members.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu bia Tiến sĩ làm bằng đá "stone slabs" và ghi tên những người đỗ đạt, không phải hoàng tộc.'),
    mcq(26, 'Reading: True/False', 'Q26: The Temple of Literature appears on Vietnamese currency.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu hình ảnh Văn Miếu xuất hiện trên mặt sau tờ 100,000 đồng.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "scholars" in the passage is closest in meaning to:', 'A. soldiers', 'B. students and learned people', 'C. farmers', 'D. merchants', 'B', '"Scholars" = học giả, những người học rộng. Phù hợp ngữ cảnh những người đến học tại Văn Miếu.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main idea of this passage?', 'A. How to get good luck on exams', 'B. The Temple of Literature as a historical and educational symbol of Vietnam', 'C. The architecture of ancient Hanoi', 'D. The life of Emperor Ly Thanh Tong', 'B', 'Bài đọc giới thiệu Văn Miếu - Quốc Tử Giám như một biểu tượng lịch sử và giáo dục của Việt Nam.'),
    textQ(29, 'Word Form', 'The ____ of the ancient tombs provided valuable insights into the lives of past rulers. (DISCOVER)', 'discovery', 'Cần danh từ sau mạo từ "The". "Discover" (động từ) -> "discovery" (danh từ).'),
    textQ(30, 'Word Form', 'The Imperial City of Hue is ____ significant to Vietnamese culture and identity. (HISTORY)', 'historically', 'Cần trạng từ bổ nghĩa cho tính từ "significant". "History" -> "historical" -> "historically".'),
    textQ(31, 'Word Form', 'Many ____ buildings in Hoi An have been carefully preserved for over 300 years. (ANTIQUITY)', 'antique', 'Cần tính từ bổ nghĩa cho "buildings". "Antiquity" (danh từ) -> "antique" (tính từ) = cổ kính.'),
    textQ(32, 'Word Form', 'Visitors are fascinated by the ____ beauty of the ancient pagodas in the complex. (ARCHITECT)', 'architectural', 'Cần tính từ bổ nghĩa cho "beauty". "Architect" -> "architectural" (thuộc kiến trúc).'),
    textQ(33, 'Word Form', 'The ____ of the old town has attracted filmmakers from around the world. (CHARMING)', 'charm', 'Cần danh từ sau mạo từ "The". "Charming" (tính từ) -> "charm" (danh từ).'),
    textQ(34, 'Word Form', 'The local government is working to ____ the ancient citadel before it deteriorates further. (RESTORATION)', 'restore', 'Cần động từ trong cấu trúc "to V". "Restoration" -> "restore" (trùng tu).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "preserve /prɪˈzɜːrv/ verb: to keep something in its original state or prevent it from being damaged or destroyed." Complete: We must do everything we can to ____ our cultural heritage for future generations.', 'preserve', '"Preserve" (bảo tồn) phù hợp với định nghĩa và ngữ cảnh bảo vệ di sản văn hóa.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "construct /kənˈstrʌkt/ verb: to build something, especially a large structure." Complete: It took over 20 years to ____ the Great Wall of China during the Ming Dynasty.', 'construct', '"Construct" (xây dựng) phù hợp với định nghĩa về xây dựng công trình lớn.'),
    textQ(37, 'Sentence Transformation', 'People built the One Pillar Pagoda during the reign of Emperor Ly Thai Tong in 1049. (WAS)\n→ The One Pillar Pagoda ____ during the reign of Emperor Ly Thai Tong in 1049.', 'was built', 'Câu bị động thì quá khứ đơn: S (số ít) + was + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'No other heritage site in Vietnam is as visited as Ha Long Bay. (MOST)\n→ Ha Long Bay is ____ heritage site in Vietnam.', 'the most visited', 'So sánh nhất: "No other...is as + adj + as" → "This is the + adj-est/most adj + N".'),
    textQ(39, 'Sentence Transformation', 'They have not restored the ancient gate since the war ended. (LAST)\n→ The ____ was when the war ended.', 'last time the ancient gate was restored', '"S + have/has not V3/ed + since..." → "The last time + S + was V3/ed + was...".'),
    textQ(40, 'Sentence Transformation', '"Take off your shoes before entering the temple," the guide said to the tourists. (ASKED)\n→ The guide ____ off their shoes before entering the temple.', 'asked the tourists to take', '"V + O," S said → "S + asked + O + to V".'),
  ];

  seedSingleTest_({
    test_id: '21',
    title: 'Đề thi thử số 21',
    description: 'Chủ đề Lịch sử & Di tích - Đọc hiểu về di sản văn hóa Việt Nam',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 22: Chủ đề Thiên nhiên & Động vật - Từ vựng về bảo tồn và sinh thái
// ======================================================================
function seedQuestionsTest22_() {
  var clozePassage = [
    "Biodiversity is essential (17)____ maintaining the balance of ecosystems on our planet. Unfortunately, many animal",
    "and plant species are facing (18)____ due to habitat loss, poaching, and climate change. Conservation efforts",
    "around the world aim (19)____ protect endangered species and restore their natural habitats. National parks and",
    "wildlife reserves play a crucial (20)____ in providing safe environments for threatened species. In Vietnam,",
    "organizations such as the Asian Turtle Program work tirelessly to protect endangered turtles (21)____ illegal",
    "trade. There are many things that ordinary people can do to help, (22)____ donating to conservation charities and",
    "reducing the use of products that harm wildlife."
  ].join('\n\n');

  var readingPassage = [
    "Cuc Phuong National Park, established in 1962, is Vietnam's first and largest national park. Located about 120",
    "kilometers southwest of Hanoi, it covers an area of over 22,000 hectares across three provinces. The park is home",
    "to an incredible diversity of flora and fauna, including over 2,000 plant species, 135 mammal species, and more",
    "than 300 bird species. Some of the rare animals found in the park include the Delacour's langur, the Owston's",
    "civet, and the clouded leopard. Cuc Phuong also runs several successful conservation programs, including a primate",
    "rescue center that rehabilitates endangered monkeys and apes rescued from the illegal wildlife trade. The park",
    "offers visitors opportunities for trekking, birdwatching, and learning about conservation. However, like many",
    "protected areas, Cuc Phuong faces challenges from illegal logging and hunting. Park rangers work around the clock",
    "to safeguard this natural treasure, but they need more support from both the government and the public."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -s/-es endings', 'Choose the word whose underlined part is pronounced differently: A. species B. plants C. animals D. reserves', 'A. species /siːz/', 'B. plants /s/', 'C. animals /z/', 'D. reserves /z/', 'B', '"Plants" có âm cuối /t/ vô thanh nên -s đọc /s/, các từ còn lại đọc /z/ hoặc /iːz/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. wildlife B. climate C. primate D. habitat', 'A. wildlife /aɪ/', 'B. climate /aɪ/', 'C. primate /aɪ/', 'D. habitat /æ/', 'D', '"Habitat" có âm "a" đầu là /æ/, các từ còn lại có âm /aɪ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. forest B. species C. protect D. nature', 'A. forest', 'B. species', 'C. protect', 'D. nature', 'C', '"Protect" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. habitat B. extinction C. national D. animal', 'A. habitat', 'B. extinction', 'C. national', 'D. animal', 'B', '"Extinction" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The natural ____ of many wild animals has been destroyed by deforestation and urbanization.', 'A. habitat', 'B. habit', 'C. custom', 'D. cuisine', 'A', '"Habitat" (môi trường sống) là thuật ngữ sinh thái học chính xác.'),
    mcq(6, 'Grammar: Tenses', 'Scientists ____ the behavior of rare langurs in Cuc Phuong National Park for over 20 years.', 'A. studied', 'B. have been studying', 'C. study', 'D. were studying', 'B', 'Thì hiện tại hoàn thành tiếp diễn với "for over 20 years" nhấn mạnh hành động liên tục.'),
    mcq(7, 'Phrasal Verbs', 'The wildlife rescue center ____ orphaned and injured animals before releasing them back into the wild.', 'A. looks after', 'B. looks down', 'C. looks for', 'D. looks into', 'A', '"Look after" = chăm sóc. Phù hợp ngữ cảnh trung tâm cứu hộ chăm sóc động vật.'),
    mcq(8, 'Prepositions', 'Many species of frogs are ____ threat due to pollution and the destruction of wetlands.', 'A. under', 'B. in', 'C. at', 'D. on', 'A', '"Under threat" = đang bị đe dọa. Là cụm giới từ cố định.'),
    mcq(9, 'Grammar: Conditional', 'If we ____ immediate action to protect wildlife, many species will become extinct within decades.', 'A. do not take', 'B. did not take', 'C. will not take', 'D. would not take', 'A', 'Câu điều kiện loại 1: If + S + V (HTĐ), S + will + V.'),
    mcq(10, 'Grammar: Passive Voice', 'Endangered animals ____ in national parks and wildlife reserves to ensure their survival.', 'A. protect', 'B. are protected', 'C. protected', 'D. have protected', 'B', 'Câu bị động thì hiện tại đơn: S (số nhiều) + are + V3/ed.'),
    mcq(11, 'Vocabulary: Word Choice', 'Cuc Phuong National Park has rich ____, with thousands of different plant and animal species.', 'A. biodiversity', 'B. geography', 'C. history', 'D. industry', 'A', '"Biodiversity" (đa dạng sinh học) là thuật ngữ chuyên ngành chính xác.'),
    mcq(12, 'Communication', 'Student: "What can I do to help protect endangered animals?"\nTeacher: "____"', 'A. You can start by learning about them and supporting wildlife conservation organizations.', 'B. There is nothing you can do about it.', 'C. Endangered animals live far away from here.', 'D. You should not worry about animals.', 'A', 'Lời khuyên thiết thực và tích cực về cách học sinh có thể giúp bảo vệ động vật.'),
    mcq(13, 'Communication', 'Visitor: "Can I feed the monkeys in the national park?"\nPark ranger: "____"', 'A. No, feeding wild animals is strictly prohibited as it harms their health and behavior.', 'B. Yes, they love bananas and peanuts.', 'C. Only if you have leftover food.', 'D. The monkeys are not here right now.', 'A', 'Kiểm lâm giải thích lý do không được cho động vật hoang dã ăn, vừa nghiêm túc vừa mang tính giáo dục.'),
    mcq(14, 'Public Signs', 'A sign in a national park showing a crossed-out campfire means:', 'A. Camping area ahead', 'B. No campfires — fire hazard', 'C. Barbecue station', 'D. Firewood for sale', 'B', 'Biển báo lửa trại bị gạch chéo cảnh báo cấm đốt lửa để phòng cháy rừng.'),
    mcq(15, 'Public Signs', 'What does a sign showing a person throwing trash into a bin at a nature reserve mean?', 'A. Do not litter — use the bin', 'B. Trash collection schedule', 'C. Recycling center location', 'D. Garbage truck entrance', 'A', 'Biển báo người bỏ rác vào thùng nhắc nhở không xả rác, hãy sử dụng thùng rác.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The government has decided ____ a new wildlife sanctuary in the Central Highlands.', 'A. establish', 'B. to establish', 'C. establishing', 'D. established', 'B', '"Decide + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Biodiversity is essential (17)____ maintaining the balance of ecosystems..."', 'A. for', 'B. to', 'C. in', 'D. with', 'A', '"Essential for + N/V-ing" = cần thiết cho việc gì.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...are facing (18)____ due to habitat loss, poaching, and climate change."', 'A. extinction', 'B. extinct', 'C. extinctive', 'D. extinguish', 'A', '"Face extinction" = đối mặt với nguy cơ tuyệt chủng. Cần danh từ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "Conservation efforts around the world aim (19)____ protect endangered species..."', 'A. to', 'B. at', 'C. for', 'D. in', 'A', '"Aim to V" = nhằm mục đích làm gì. Cấu trúc cố định.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...wildlife reserves play a crucial (20)____ in providing safe environments..."', 'A. role', 'B. part', 'C. function', 'D. job', 'A', '"Play a crucial role in + N/V-ing" = đóng vai trò quan trọng trong.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...protect endangered turtles (21)____ illegal trade."', 'A. from', 'B. against', 'C. off', 'D. away', 'A', '"Protect...from + N" = bảo vệ...khỏi cái gì.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...things that ordinary people can do to help, (22)____ donating to conservation..."', 'A. such as', 'B. so that', 'C. because of', 'D. in case', 'A', '"Such as" = ví dụ như, dùng để liệt kê các hành động cụ thể.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: When was Cuc Phuong National Park established?', 'A. In 1952', 'B. In 1962', 'C. In 1972', 'D. In 1982', 'B', 'Đoạn văn nêu: "established in 1962".'),
    mcq(24, 'Reading: True/False', 'Q24: Cuc Phuong is the smallest national park in Vietnam.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu đây là "Vietnam\'s first and largest national park".'),
    mcq(25, 'Reading: True/False', 'Q25: Cuc Phuong has a primate rescue center that rehabilitates endangered monkeys and apes.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "a primate rescue center that rehabilitates endangered monkeys and apes".'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, Cuc Phuong faces no challenges or threats at all.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu công viên đối mặt với "illegal logging and hunting".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "rehabilitates" in the passage most likely means:', 'A. captures wild animals for zoos', 'B. helps animals recover and return to normal life', 'C. moves animals to other countries', 'D. studies animals in laboratories', 'B', '"Rehabilitate" = phục hồi chức năng, giúp động vật hồi phục và trở lại cuộc sống bình thường.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of this passage?', 'A. To promote tourism to Cuc Phuong National Park', 'B. To describe Cuc Phuong National Park and the importance of conservation efforts', 'C. To list all animal species in Vietnam', 'D. To argue that national parks should be closed to visitors', 'B', 'Bài đọc mô tả Vườn quốc gia Cúc Phương và nhấn mạnh tầm quan trọng của công tác bảo tồn.'),
    textQ(29, 'Word Form', 'The ____ of coral reefs is a serious concern for marine biologists worldwide. (DESTROY)', 'destruction', 'Cần danh từ sau mạo từ "The". "Destroy" (động từ) -> "destruction" (danh từ).'),
    textQ(30, 'Word Form', 'Tigers are ____ endangered, with fewer than 200 individuals left in the wild in Vietnam. (CRITICAL)', 'critically', 'Cần trạng từ bổ nghĩa cho tính từ "endangered". "Critical" -> "critically".'),
    textQ(31, 'Word Form', 'There are many ____ programs working to save the saola, one of the rarest animals in the world. (CONSERVE)', 'conservation', 'Cần danh từ bổ nghĩa cho "programs". "Conserve" -> "conservation" (bảo tồn).'),
    textQ(32, 'Word Form', 'The park rangers work ____ to protect wildlife from illegal poachers day and night. (TIRE)', 'tirelessly', 'Cần trạng từ bổ nghĩa cho "work". "Tire" -> "tireless" -> "tirelessly".'),
    textQ(33, 'Word Form', 'The ____ of the forest ecosystem depends on the survival of all species within it. (STABLE)', 'stability', 'Cần danh từ sau mạo từ "The". "Stable" (tính từ) -> "stability" (danh từ).'),
    textQ(34, 'Word Form', 'Many ____ plants in Cuc Phuong are used in traditional medicine. (VALUE)', 'valuable', 'Cần tính từ bổ nghĩa cho "plants". "Value" -> "valuable" (có giá trị).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "poach /pəʊtʃ/ verb: to illegally hunt or catch animals on land that is not one\'s own." Complete: Criminals who ____ elephants for their ivory tusks face heavy fines and prison sentences.', 'poach', '"Poach" (săn trộm) phù hợp với định nghĩa và ngữ cảnh săn voi lấy ngà.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "migrate /maɪˈɡreɪt/ verb: to move from one habitat or region to another, typically according to the seasons." Complete: Many bird species ____ south during the winter to find warmer climates and more food.', 'migrate', '"Migrate" (di cư) phù hợp với định nghĩa và ngữ cảnh chim di cư về phương nam.'),
    textQ(37, 'Sentence Transformation', 'Governments should enforce stricter laws to protect wildlife. (ENFORCED)\n→ Stricter laws ____ to protect wildlife by governments.', 'should be enforced', 'Câu bị động với động từ khuyết thiếu: S + should + be + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'The number of wild tigers has decreased dramatically in the past century. (DECREASE)\n→ There ____ in the number of wild tigers in the past century.', 'has been a dramatic decrease', '"S + has decreased dramatically" → "There has been a dramatic decrease in + N".'),
    textQ(39, 'Sentence Transformation', 'Many animals lose their homes because forests are cut down for farming. (LOSS)\n→ Many animals suffer from ____ due to deforestation for farming.', 'the loss of their homes', '"Lose their homes" → "the loss of their homes". Chuyển từ động từ sang danh từ.'),
    textQ(40, 'Sentence Transformation', '"Stay on the trails and do not disturb the wildlife," the guide reminded the hikers. (REMINDED)\n→ The guide ____ on the trails and not to disturb the wildlife.', 'reminded the hikers to stay', '"V," S reminded O → "S + reminded + O + to V / not to V".'),
  ];

  seedSingleTest_({
    test_id: '22',
    title: 'Đề thi thử số 22',
    description: 'Chủ đề Thiên nhiên & Động vật - Từ vựng về bảo tồn và sinh thái',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 23: Chủ đề Sách & Văn hóa đọc - Đọc hiểu về thói quen đọc sách
// ======================================================================
function seedQuestionsTest23_() {
  var clozePassage = [
    "Reading is one of the most beneficial habits a person (17)____ develop. Studies have shown that regular reading",
    "improves vocabulary, enhances concentration, and reduces (18)____. Despite these benefits, reading habits among",
    "young people have declined in the digital (19)____. Many teenagers prefer spending their free time on smartphones",
    "and video games (20)____ than reading books. Parents and teachers can encourage reading (21)____ creating a",
    "comfortable reading environment and letting children choose books that interest them. Public libraries also play",
    "an important role (22)____ making books accessible to everyone, especially in underprivileged communities."
  ].join('\n\n');

  var readingPassage = [
    "The annual Ho Chi Minh City Book Fair is one of the largest literary events in Vietnam, attracting hundreds of",
    "publishers and thousands of book lovers each year. The fair features a wide range of books, from Vietnamese",
    "literature and translated foreign works to academic textbooks and children's stories. In recent years, the fair",
    "has also embraced digital reading, with e-book platforms and audiobook services setting up interactive booths.",
    "Visitors can meet their favorite authors, attend book signings, and participate in writing workshops and panel",
    "discussions. The event is particularly popular among young readers and families. Organizers say the fair's mission",
    "is not just to sell books but to promote a culture of reading in Vietnamese society. According to a survey, the",
    "average Vietnamese person reads about four books per year — a number that organizers hope to increase through",
    "events like the book fair. 'Reading opens the mind to new ideas and perspectives,' said one author at the event.",
    "'In a world of short videos and instant messages, taking time to read a book is a form of quiet rebellion.'"
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. reading B. teacher C. create D. great', 'A. reading /iː/', 'B. teacher /iː/', 'C. create /iː/', 'D. great /eɪ/', 'D', '"Great" có âm "ea" đọc là /eɪ/, các từ còn lại đọc là /iː/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. chapter B. character C. children D. choose', 'A. chapter /tʃ/', 'B. character /k/', 'C. children /tʃ/', 'D. choose /tʃ/', 'B', '"Character" có âm "ch" đọc là /k/, các từ còn lại đọc là /tʃ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. novel B. author C. essay D. review', 'A. novel', 'B. author', 'C. essay', 'D. review', 'D', '"Review" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. library B. publisher C. magazine D. literature', 'A. library', 'B. publisher', 'C. magazine', 'D. literature', 'C', '"Magazine" nhấn âm 3, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'A good book can ____ your imagination and transport you to entirely different worlds and times.', 'A. stimulate', 'B. reduce', 'C. prevent', 'D. destroy', 'A', '"Stimulate imagination" (kích thích trí tưởng tượng) là lợi ích của việc đọc sách.'),
    mcq(6, 'Grammar: Tenses', 'She ____ six novels this year and plans to read at least four more before December.', 'A. reads', 'B. has read', 'C. read', 'D. is reading', 'B', 'Thì hiện tại hoàn thành với "this year" — khoảng thời gian chưa kết thúc.'),
    mcq(7, 'Phrasal Verbs', 'I could not ____ the book because the plot was too confusing and the characters were uninteresting.', 'A. get through', 'B. take off', 'C. put up', 'D. give away', 'A', '"Get through" = đọc hết, hoàn thành (một cuốn sách). Phù hợp ngữ cảnh không thể đọc hết.'),
    mcq(8, 'Prepositions', 'Reading books is a great way to escape ____ the stress and pressure of daily life.', 'A. from', 'B. of', 'C. for', 'D. at', 'A', '"Escape from + N" = thoát khỏi điều gì.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish I ____ more time to read all the books on my shelf that I have not opened yet.', 'A. have', 'B. had', 'C. will have', 'D. having', 'B', 'Câu ước ở hiện tại: S + wish + S + V (quá khứ đơn).'),
    mcq(10, 'Grammar: Relative Clause', 'The library ____ I spend most weekends has an impressive collection of classic novels.', 'A. where', 'B. which', 'C. who', 'D. what', 'A', '"Where" là trạng từ quan hệ thay thế cho cụm chỉ nơi chốn "the library".'),
    mcq(11, 'Vocabulary: Word Choice', 'The Harry Potter series, ____ by J.K. Rowling, has sold over 500 million copies worldwide.', 'A. written', 'B. wrote', 'C. writing', 'D. writes', 'A', '"Written by" = được viết bởi. Rút gọn mệnh đề bị động.'),
    mcq(12, 'Communication', 'Nam: "What kind of books do you enjoy reading?"\nLan: "____"', 'A. I am really into mystery and detective novels.', 'B. I go to the library twice a week.', 'C. Reading is a waste of time.', 'D. My mother bought me a new book yesterday.', 'A', 'Trả lời về thể loại sách yêu thích (trinh thám) phù hợp với câu hỏi.'),
    mcq(13, 'Communication', 'Friend: "Can I borrow your copy of The Great Gatsby?"\nYou: "____"', 'A. Sure, just make sure to return it in good condition when you are done.', 'B. No, I do not read books.', 'C. What is the book about?', 'D. The library is closed on Sundays.', 'A', 'Đồng ý cho mượn sách kèm lời nhắn giữ gìn cẩn thận là phản hồi lịch sự.'),
    mcq(14, 'Public Signs', 'A sign showing a person reading a book with "Quiet Zone" below it means:', 'A. Bookstore ahead', 'B. Silent area — no talking or noise', 'C. Reading club meeting point', 'D. Free book exchange', 'B', 'Biển báo "Quiet Zone" với hình người đọc sách chỉ khu vực yên tĩnh, không gây ồn ào.'),
    mcq(15, 'Public Signs', 'What does a sign with an open book and a "+" symbol usually mean at a café?', 'A. Bookstore section', 'B. Book crossing / free book exchange shelf', 'C. Library checkout desk', 'D. Study room reservation', 'B', 'Biển báo sách mở với dấu + thường chỉ điểm trao đổi sách miễn phí (book crossing).'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She finished ____ the novel late at night and could not stop thinking about the ending.', 'A. read', 'B. to read', 'C. reading', 'D. having read', 'C', '"Finish + V-ing" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Reading is one of the most beneficial habits a person (17)____ develop."', 'A. can', 'B. must', 'C. should', 'D. need', 'A', '"Can" diễn tả khả năng. "A person can develop" = một người có thể phát triển.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...improves vocabulary, enhances concentration, and reduces (18)____."', 'A. stress', 'B. happiness', 'C. excitement', 'D. energy', 'A', '"Reduces stress" = giảm căng thẳng, là một lợi ích đã được chứng minh của việc đọc sách.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...have declined in the digital (19)____."', 'A. age', 'B. time', 'C. period', 'D. generation', 'A', '"In the digital age" = trong thời đại kỹ thuật số. Là cụm từ cố định.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...prefer spending their free time on smartphones and video games (20)____ than reading books."', 'A. rather', 'B. better', 'C. more', 'D. sooner', 'A', '"Rather than" = hơn là. Cấu trúc so sánh lựa chọn.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "Parents and teachers can encourage reading (21)____ creating a comfortable..."', 'A. by', 'B. with', 'C. for', 'D. in', 'A', '"By + V-ing" = bằng cách làm gì. Chỉ phương thức.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...play an important role (22)____ making books accessible to everyone..."', 'A. in', 'B. on', 'C. at', 'D. for', 'A', '"Play a role in + N/V-ing" = đóng vai trò trong việc gì.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, what is one of the missions of the Ho Chi Minh City Book Fair?', 'A. To sell as many books as possible at high prices', 'B. To promote a culture of reading in Vietnamese society', 'C. To replace all physical books with e-books', 'D. To discourage young people from reading', 'B', 'Đoạn văn nêu sứ mệnh của hội sách là "to promote a culture of reading in Vietnamese society".'),
    mcq(24, 'Reading: True/False', 'Q24: The book fair only features Vietnamese literature and does not include foreign books.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu hội sách trưng bày "translated foreign works" bên cạnh văn học Việt Nam.'),
    mcq(25, 'Reading: True/False', 'Q25: The average Vietnamese person reads about ten books per year.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu trung bình người Việt đọc "about four books per year", không phải mười.'),
    mcq(26, 'Reading: True/False', 'Q26: The book fair includes activities such as author meet-and-greets and writing workshops.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn liệt kê: "meet their favorite authors, attend book signings, and participate in writing workshops".'),
    mcq(27, 'Reading: MCQ', 'Q27: The phrase "quiet rebellion" used by the author in the passage suggests that reading books today is:', 'A. An illegal activity', 'B. An act of resisting the trend of short attention spans and digital distractions', 'C. Something only old people do', 'D. A way to protest against the government', 'B', '"Quiet rebellion" (nổi loạn thầm lặng) ám chỉ việc đọc sách là hành động đi ngược xu hướng nội dung ngắn và xao lãng kỹ thuật số.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of this passage?', 'A. The Ho Chi Minh City Book Fair is only for publishers to make money', 'B. The book fair plays an important role in promoting reading culture in Vietnam', 'C. Vietnamese people read more than anyone else in the world', 'D. E-books will soon replace all printed books in Vietnam', 'B', 'Bài đọc nhấn mạnh vai trò của hội sách trong việc thúc đẩy văn hóa đọc tại Việt Nam.'),
    textQ(29, 'Word Form', 'Reading books regularly can help improve your ____ and critical thinking skills. (CREATE)', 'creativity', 'Cần danh từ sau tính từ sở hữu "your". "Create" (động từ) -> "creativity" (danh từ).'),
    textQ(30, 'Word Form', 'The author was ____ happy when her first novel became a bestseller within a month. (EXTREME)', 'extremely', 'Cần trạng từ bổ nghĩa cho tính từ "happy". "Extreme" -> "extremely".'),
    textQ(31, 'Word Form', 'Many classic works of ____ have been adapted into successful movies and TV series. (LITERATE)', 'literature', 'Cần danh từ sau giới từ "of". "Literate" (tính từ) -> "literature" (danh từ).'),
    textQ(32, 'Word Form', 'A good ____ can make any story come alive through their expressive reading style. (NARRATE)', 'narrator', 'Cần danh từ chỉ người. "Narrate" (động từ) -> "narrator" (người kể chuyện).'),
    textQ(33, 'Word Form', 'The ____ of e-books has made reading more accessible to people in remote areas. (POPULAR)', 'popularity', 'Cần danh từ sau mạo từ "The". "Popular" (tính từ) -> "popularity" (danh từ).'),
    textQ(34, 'Word Form', 'Reading is an ____ part of a child\'s cognitive and emotional development. (ESSENCE)', 'essential', 'Cần tính từ bổ nghĩa cho "part". "Essence" -> "essential".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "publish /ˈpʌblɪʃ/ verb: to prepare and issue a book, journal, or piece of music for public sale." Complete: The company plans to ____ her second novel in the spring of next year.', 'publish', '"Publish" (xuất bản) phù hợp với định nghĩa và ngữ cảnh công ty xuất bản tiểu thuyết.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "skim /skɪm/ verb: to read something quickly to get the main ideas without reading every word." Complete: Before the exam, she ____ through her notes to review the key concepts one last time.', 'skimmed', '"Skim through" (đọc lướt qua) phù hợp với định nghĩa và ngữ cảnh ôn bài trước kỳ thi.'),
    textQ(37, 'Sentence Transformation', 'She has not read a novel since she started her new job last year. (LAST)\n→ The ____ was when she started her new job last year.', 'last time she read a novel', '"S + have/has not V3/ed + since..." → "The last time + S + V (QKĐ) + was...".'),
    textQ(38, 'Sentence Transformation', '"You should read more books to expand your vocabulary," the teacher told us. (ADVISED)\n→ The teacher ____ more books to expand our vocabulary.', 'advised us to read', '"You should V," S told O → "S + advised + O + to V".'),
    textQ(39, 'Sentence Transformation', 'The novel is so interesting that I cannot put it down. (SUCH)\n→ It is ____ that I cannot put it down.', 'such an interesting novel', '"So + adj + that" → "Such + (a/an) + adj + N + that".'),
    textQ(40, 'Sentence Transformation', 'People say that reading fiction improves empathy and emotional intelligence. (SAID)\n→ Reading fiction ____ empathy and emotional intelligence.', 'is said to improve', '"People say that S + V" → "S + is/are said + to V".'),
  ];

  seedSingleTest_({
    test_id: '23',
    title: 'Đề thi thử số 23',
    description: 'Chủ đề Sách & Văn hóa đọc - Đọc hiểu về thói quen đọc sách',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 24: Chủ đề Năng lượng tái tạo - Câu điều kiện và câu bị động
// ======================================================================
function seedQuestionsTest24_() {
  var clozePassage = [
    "Renewable energy sources (17)____ as solar, wind, and hydroelectric power are becoming increasingly important",
    "in the fight against climate change. Unlike fossil fuels, (18)____ are finite and polluting, renewable energy",
    "is clean and virtually inexhaustible. Many countries are investing (19)____ in renewable energy infrastructure",
    "to reduce their dependence on oil and coal. Solar panels, which convert sunlight (20)____ electricity, are now",
    "commonly installed on rooftops across Vietnam. If the world transitions fully to renewable energy, carbon",
    "emissions (21)____ be reduced by more than 70 percent. However, challenges such as energy storage and high",
    "initial costs still need to (22)____ addressed before renewables can fully replace fossil fuels."
  ].join('\n\n');

  var readingPassage = [
    "Solar energy is one of the most promising renewable energy sources available today. The sun provides more energy",
    "to the Earth in one hour than the entire world consumes in a year. Solar panels, also known as photovoltaic cells,",
    "capture sunlight and convert it directly into electricity. In Vietnam, the government has introduced policies",
    "encouraging both households and businesses to install solar power systems. The feed-in tariff program, which",
    "allows solar panel owners to sell excess electricity back to the national grid, has been particularly successful.",
    "As a result, Vietnam has become one of the leading countries in Southeast Asia for solar energy adoption. However,",
    "solar energy also has its limitations. It is intermittent — solar panels only generate electricity when the sun",
    "is shining. This means energy storage solutions, such as batteries, are essential for making solar power a reliable",
    "energy source around the clock. Researchers are working on developing more efficient and affordable batteries to",
    "address this challenge. If these storage problems can be solved, solar energy could power a much larger share of",
    "the world's energy needs in the coming decades."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. solar B. power C. coal D. fossil', 'A. solar /əʊ/', 'B. power /aʊ/', 'C. coal /əʊ/', 'D. fossil /ɒ/', 'B', '"Power" có âm "ow" đọc là /aʊ/, các từ còn lại có âm /əʊ/ hoặc /ɒ/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. energy B. generate C. storage D. panel', 'A. energy /dʒ/', 'B. generate /dʒ/', 'C. storage /dʒ/', 'D. panel /p/', 'D', '"Panel" có âm "p", các từ còn lại có âm /dʒ/ (g). Khác biệt phụ âm đầu.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. solar B. fossil C. turbine D. supply', 'A. solar', 'B. fossil', 'C. turbine', 'D. supply', 'D', '"Supply" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. energy B. electric C. battery D. company', 'A. energy', 'B. electric', 'C. battery', 'D. company', 'B', '"Electric" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'Wind turbines ____ the kinetic energy of the wind into mechanical power, which is then converted to electricity.', 'A. convert', 'B. consume', 'C. waste', 'D. absorb', 'A', '"Convert into" (chuyển đổi thành) là thuật ngữ chính xác về biến đổi năng lượng.'),
    mcq(6, 'Grammar: Tenses', 'By 2030, Vietnam ____ enough solar capacity to power over five million households.', 'A. will install', 'B. will have installed', 'C. installed', 'D. installs', 'B', 'Thì tương lai hoàn thành: By + thời điểm TL, S + will have V3/ed.'),
    mcq(7, 'Phrasal Verbs', 'Scientists are trying to ____ more efficient ways to store solar energy for nighttime use.', 'A. come up with', 'B. run out of', 'C. put up with', 'D. look forward to', 'A', '"Come up with" = nghĩ ra, tìm ra. "Come up with ways" = tìm ra cách.'),
    mcq(8, 'Prepositions', 'Solar panels are typically installed ____ the rooftops of buildings to capture maximum sunlight.', 'A. on', 'B. in', 'C. at', 'D. under', 'A', '"On the rooftops" = trên mái nhà. Giới từ chỉ vị trí trên bề mặt.'),
    mcq(9, 'Grammar: Conditional Type 2', 'If every household in Vietnam ____ solar panels, the country would save millions in electricity costs.', 'A. installs', 'B. installed', 'C. will install', 'D. would install', 'B', 'Câu điều kiện loại 2: If + S + V (QKĐ), S + would + V.'),
    mcq(10, 'Grammar: Passive Voice', 'The new wind farm ____ next month and will supply electricity to 50,000 homes.', 'A. will complete', 'B. will be completed', 'C. completed', 'D. is completing', 'B', 'Câu bị động thì tương lai đơn: S + will be + V3/ed.'),
    mcq(11, 'Vocabulary: Word Choice', 'One of the main ____ of renewable energy is that it does not produce harmful greenhouse gases.', 'A. advantages', 'B. disadvantages', 'C. problems', 'D. difficulties', 'A', '"Advantages" (lợi thế, ưu điểm) phù hợp khi nói về mặt tích cực của năng lượng tái tạo.'),
    mcq(12, 'Communication', 'Neighbor: "Are those solar panels on your roof? How do they work for you?"\nYou: "____"', 'A. They have reduced our electricity bill by about 40%. It was a great investment.', 'B. I do not like the color of the panels.', 'C. Solar panels are too expensive for most people.', 'D. I wish I had never installed them.', 'A', 'Chia sẻ trải nghiệm tích cực về pin mặt trời, trả lời đúng câu hỏi của hàng xóm.'),
    mcq(13, 'Communication', 'Student: "What can our school do to use more green energy?"\nPrincipal: "____"', 'A. We could install solar panels on the gymnasium roof and use the savings for educational programs.', 'B. Green energy is not important for schools.', 'C. Students should not worry about energy issues.', 'D. The school does not have any electricity problems.', 'A', 'Hiệu trưởng đưa ra đề xuất cụ thể và thiết thực về việc sử dụng năng lượng xanh.'),
    mcq(14, 'Public Signs', 'A sign showing a lightning bolt inside a green leaf means:', 'A. Danger: electric fence', 'B. Green electricity / renewable energy zone', 'C. High voltage — keep away', 'D. Electric vehicle charging only', 'B', 'Biểu tượng tia sét trong chiếc lá xanh thể hiện năng lượng xanh/năng lượng tái tạo.'),
    mcq(15, 'Public Signs', 'What does a sign with a crossed-out incandescent bulb at a store mean?', 'A. Light bulbs sold here', 'B. Incandescent bulbs are banned — use LED bulbs only', 'C. No lighting in this area', 'D. Emergency lighting only', 'B', 'Biển báo bóng đèn sợi đốt bị gạch chéo khuyến khích hoặc yêu cầu sử dụng đèn LED tiết kiệm điện.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'We need ____ more in renewable energy research if we want to combat climate change effectively.', 'A. invest', 'B. to invest', 'C. investing', 'D. invested', 'B', '"Need + to V" = cần làm gì. Cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Renewable energy sources (17)____ as solar, wind, and hydroelectric power..."', 'A. such', 'B. like', 'C. same', 'D. similar', 'A', '"Such as" = ví dụ như. Dùng để liệt kê.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "Unlike fossil fuels, (18)____ are finite and polluting..."', 'A. which', 'B. that', 'C. who', 'D. what', 'A', '"Which" thay cho "fossil fuels" trong mệnh đề quan hệ không xác định.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "Many countries are investing (19)____ in renewable energy infrastructure..."', 'A. heavily', 'B. heavy', 'C. heaviness', 'D. heavier', 'A', 'Cần trạng từ bổ nghĩa cho động từ "investing". "Heavy" -> "heavily".'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...which convert sunlight (20)____ electricity..."', 'A. into', 'B. to', 'C. for', 'D. with', 'A', '"Convert A into B" = chuyển đổi A thành B.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...carbon emissions (21)____ be reduced by more than 70 percent."', 'A. could', 'B. must', 'C. should', 'D. have', 'A', '"Could" diễn tả khả năng trong câu điều kiện.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...challenges...still need to (22)____ addressed..."', 'A. be', 'B. have', 'C. do', 'D. get', 'A', '"Need to be + V3/ed" = cần được. Cấu trúc bị động.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, how much energy does the sun provide to Earth in one hour?', 'A. Less than the world consumes in a day', 'B. More than the world consumes in a year', 'C. Exactly the same as the world consumes in a month', 'D. Twice as much as the world produces from coal', 'B', 'Đoạn văn nêu: "The sun provides more energy to the Earth in one hour than the entire world consumes in a year."'),
    mcq(24, 'Reading: True/False', 'Q24: Solar panels can only generate electricity during the daytime when the sun is shining.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "solar panels only generate electricity when the sun is shining".'),
    mcq(25, 'Reading: True/False', 'Q25: Vietnam has become one of the leading countries in Southeast Asia for solar energy adoption.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "Vietnam has become one of the leading countries in Southeast Asia for solar energy adoption."'),
    mcq(26, 'Reading: True/False', 'Q26: According to the passage, energy storage is not a problem for solar power at all.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu tính gián đoạn là hạn chế và pin lưu trữ rất cần thiết, đây là một thách thức.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "intermittent" in the passage most likely means:', 'A. constant and continuous', 'B. occurring at irregular intervals, not continuous', 'C. extremely powerful', 'D. cheap and affordable', 'B', '"Intermittent" = ngắt quãng, không liên tục. Chỉ việc pin mặt trời chỉ tạo điện khi có nắng.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the author\'s main point about solar energy?', 'A. Solar energy is too expensive and unreliable to be useful', 'B. Solar energy has great potential but requires better storage solutions', 'C. Solar panels should be banned because they are ugly', 'D. Vietnam should stop investing in solar energy immediately', 'B', 'Tác giả nhấn mạnh tiềm năng lớn của năng lượng mặt trời nhưng cần giải quyết vấn đề lưu trữ.'),
    textQ(29, 'Word Form', 'The ____ of renewable energy is crucial for a sustainable and environmentally friendly future. (ADOPT)', 'adoption', 'Cần danh từ sau mạo từ "The". "Adopt" (động từ) -> "adoption" (danh từ).'),
    textQ(30, 'Word Form', 'The government is ____ supporting the development of wind farms along the central coast. (ACT)', 'actively', 'Cần trạng từ bổ nghĩa cho động từ "supporting". "Act" -> "active" -> "actively".'),
    textQ(31, 'Word Form', 'Solar energy is an ____ source of power that will never run out. (EXHAUST)', 'inexhaustible', 'Cần tính từ có nghĩa phủ định "không cạn kiệt". "Exhaust" -> "inexhaustible".'),
    textQ(32, 'Word Form', 'The ____ of fossil fuels has caused significant damage to the environment worldwide. (BURN)', 'burning', 'Cần danh từ (danh động từ) sau mạo từ "The". "Burn" -> "burning" (sự đốt cháy).'),
    textQ(33, 'Word Form', 'Switching to renewable energy is an ____ sound decision for both the economy and the environment. (ECONOMY)', 'economically', 'Cần trạng từ bổ nghĩa cho tính từ "sound". "Economy" -> "economical" -> "economically".'),
    textQ(34, 'Word Form', 'The ____ of new battery technologies is key to making renewable energy more reliable. (DEVELOP)', 'development', 'Cần danh từ sau mạo từ "The". "Develop" -> "development".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "harness /ˈhɑːrnɪs/ verb: to control and use the natural force or power of something." Complete: Engineers are finding new ways to ____ the power of ocean waves to generate clean electricity.', 'harness', '"Harness the power" (khai thác sức mạnh) phù hợp với định nghĩa và ngữ cảnh năng lượng sóng biển.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "sustain /səˈsteɪn/ verb: to maintain something at a certain rate or level without depleting resources." Complete: We need to find ways to ____ our current standard of living without destroying the planet.', 'sustain', '"Sustain" (duy trì bền vững) phù hợp với định nghĩa về phát triển bền vững.'),
    textQ(37, 'Sentence Transformation', 'They will build a new solar power plant in Ninh Thuan province next year. (BUILT)\n→ A new solar power plant ____ in Ninh Thuan province next year.', 'will be built', 'Câu bị động thì tương lai đơn: S + will be + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'If we do not invest in renewable energy now, we will regret it later. (UNLESS)\n→ ____ in renewable energy now, we will regret it later.', 'Unless we invest', '"If...not" → "Unless". Unless + S + V (HTĐ) = If + S + do/does not + V.'),
    textQ(39, 'Sentence Transformation', 'The price of solar panels has decreased dramatically over the last ten years. (DECREASE)\n→ There ____ in the price of solar panels over the last ten years.', 'has been a dramatic decrease', '"S + has decreased dramatically" → "There has been a dramatic decrease in + N".'),
    textQ(40, 'Sentence Transformation', '"Turn off the lights when you leave the room," my father always says to me. (REMINDED)\n→ My father always ____ the lights when I leave the room.', 'reminds me to turn off', '"V," S says to O → "S + reminds + O + to V". Thì hiện tại vì "always".'),
  ];

  seedSingleTest_({
    test_id: '24',
    title: 'Đề thi thử số 24',
    description: 'Chủ đề Năng lượng tái tạo - Câu điều kiện và câu bị động',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 25: Chủ đề Đô thị thông minh - Từ vựng về thành phố tương lai
// ======================================================================
function seedQuestionsTest25_() {
  var clozePassage = [
    "Smart cities use technology and data (17)____ improve the quality of life for their residents. Sensors and",
    "Internet of Things (IoT) devices collect (18)____ information about traffic, air quality, and energy usage.",
    "This data (19)____ city authorities to make informed decisions about urban planning and resource management.",
    "Cities like Singapore, Barcelona, and Seoul are leading the way (20)____ implementing smart city solutions.",
    "However, building a smart city requires significant (21)____ and careful consideration of privacy concerns.",
    "Despite these challenges, experts believe that smart cities represent (22)____ future of urban living worldwide."
  ].join('\n\n');

  var readingPassage = [
    "Singapore is widely regarded as one of the world's leading smart cities. The city-state has implemented numerous",
    "technology-driven solutions to improve urban life. For example, its Smart Nation initiative uses sensors and cameras",
    "to monitor traffic flow in real time, adjusting traffic light timings to reduce congestion. The city also has a",
    "sophisticated water management system that recycles wastewater and collects rainwater, making the city more",
    "resilient to water shortages. In public housing estates, smart meters help residents track their electricity and",
    "water usage, encouraging conservation. Singapore's approach to urban planning integrates green spaces throughout",
    "the city, with vertical gardens on buildings and park connectors linking neighborhoods. Perhaps most impressively,",
    "Singapore has created a 'digital twin' — a virtual 3D model of the entire city that allows planners to simulate",
    "and test urban planning decisions before implementing them in the real world. Other cities around the globe are",
    "now studying Singapore's model as they develop their own smart city strategies."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. monitored B. integrated C. implemented D. adjusted', 'A. monitored /d/', 'B. integrated /ɪd/', 'C. implemented /ɪd/', 'D. adjusted /ɪd/', 'A', '"Monitored" có âm cuối /r/ hữu thanh nên -ed đọc /d/, các từ còn lại có âm cuối /t/ nên -ed đọc /ɪd/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. smart B. traffic C. data D. waste', 'A. smart /ɑː/', 'B. traffic /æ/', 'C. data /eɪ/', 'D. waste /eɪ/', 'A', '"Smart" có âm "ar" đọc là /ɑː/, các từ còn lại có âm /æ/ hoặc /eɪ/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. sensor B. device C. data D. traffic', 'A. sensor', 'B. device', 'C. data', 'D. traffic', 'B', '"Device" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. resident B. solution C. quality D. management', 'A. resident', 'B. solution', 'C. quality', 'D. management', 'B', '"Solution" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'Smart sensors can ____ air pollution levels in real time and alert residents when they reach dangerous levels.', 'A. monitor', 'B. ignore', 'C. increase', 'D. hide', 'A', '"Monitor" (giám sát, theo dõi) là chức năng chính của cảm biến trong đô thị thông minh.'),
    mcq(6, 'Grammar: Tenses', 'By 2030, many cities around the world ____ smart traffic systems to reduce congestion.', 'A. will install', 'B. will have installed', 'C. installed', 'D. are installing', 'B', 'Tương lai hoàn thành: By + mốc TL, S + will have V3/ed.'),
    mcq(7, 'Phrasal Verbs', 'The city plans to ____ a new smart parking system that guides drivers to available spaces.', 'A. set up', 'B. break down', 'C. put off', 'D. give away', 'A', '"Set up" = thiết lập, lắp đặt. "Set up a new system" = thiết lập hệ thống mới.'),
    mcq(8, 'Prepositions', 'Residents in smart cities have access ____ real-time information about public transportation and traffic.', 'A. to', 'B. with', 'C. for', 'D. at', 'A', '"Have access to + N" = có quyền truy cập vào. Cấu trúc cố định.'),
    mcq(9, 'Grammar: Conditional Type 1', 'If the city ____ smart sensors, it will be able to detect water leaks before they become major problems.', 'A. installs', 'B. installed', 'C. will install', 'D. would install', 'A', 'Câu điều kiện loại 1: If + S + V (HTĐ), S + will + V.'),
    mcq(10, 'Grammar: Passive Voice', 'In smart cities, data from sensors ____ to optimize energy consumption and reduce waste.', 'A. is used', 'B. uses', 'C. used', 'D. are using', 'A', 'Câu bị động thì hiện tại đơn: S (không đếm được "data") + is + V3/ed.'),
    mcq(11, 'Vocabulary: Word Choice', 'The city\'s ____ system uses recycled water for irrigation and industrial purposes, saving millions of liters each year.', 'A. water management', 'B. traffic control', 'C. waste collection', 'D. public transport', 'A', '"Water management system" (hệ thống quản lý nước) phù hợp với ngữ cảnh tái chế nước.'),
    mcq(12, 'Communication', 'Citizen: "How will the new smart traffic lights help our neighborhood?"\nOfficial: "____"', 'A. They will adjust in real time to reduce congestion and shorten your commute.', 'B. Smart traffic lights are too expensive for this city.', 'C. I do not know anything about the traffic system.', 'D. You should just use a bicycle instead of driving.', 'A', 'Giải thích lợi ích cụ thể của đèn giao thông thông minh, trả lời đúng câu hỏi của người dân.'),
    mcq(13, 'Communication', 'Visitor: "What makes your city a smart city?"\nGuide: "____"', 'A. We use technology and data to improve everything from public transport to waste management.', 'B. Our city has many tall buildings and shopping malls.', 'C. I am not sure what a smart city means exactly.', 'D. The city was founded over 300 years ago.', 'A', 'Giải thích khái niệm thành phố thông minh với các ví dụ cụ thể, phù hợp câu hỏi.'),
    mcq(14, 'Public Signs', 'A sign showing a QR code at a bus stop most likely means:', 'A. Scan here for bus schedule and real-time arrival information', 'B. Do not use phones at the bus stop', 'C. Bus tickets sold here only', 'D. Taxi pickup point', 'A', 'Mã QR ở trạm xe buýt thường dùng để quét lấy thông tin lịch trình và thời gian đến theo thời gian thực.'),
    mcq(15, 'Public Signs', 'What does a sign with a Wi-Fi symbol and the text "Free Public Wi-Fi" at a park mean?', 'A. Private network — do not connect', 'B. Free wireless internet access for everyone', 'C. Wi-Fi equipment storage area', 'D. Internet café nearby', 'B', 'Biển báo "Free Public Wi-Fi" cho biết khu vực có internet không dây miễn phí cho công chúng.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The city council decided ____ in smart waste bins that alert collectors when they are full.', 'A. invest', 'B. to invest', 'C. investing', 'D. invested', 'B', '"Decide + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Smart cities use technology and data (17)____ improve the quality of life..."', 'A. to', 'B. for', 'C. in', 'D. at', 'A', '"Use...to V" = sử dụng...để làm gì. Cấu trúc chỉ mục đích.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...collect (18)____ information about traffic, air quality, and energy usage."', 'A. real-time', 'B. real', 'C. time', 'D. timely', 'A', '"Real-time information" = thông tin theo thời gian thực. Là thuật ngữ công nghệ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "This data (19)____ city authorities to make informed decisions..."', 'A. allows', 'B. allows to', 'C. allowing', 'D. is allowing', 'A', '"Data" (không đếm được) + động từ số ít. "Allows + O + to V".'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...are leading the way (20)____ implementing smart city solutions."', 'A. in', 'B. on', 'C. at', 'D. for', 'A', '"Lead the way in + N/V-ing" = dẫn đầu trong lĩnh vực gì.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...requires significant (21)____ and careful consideration..."', 'A. investment', 'B. invest', 'C. investor', 'D. investing', 'A', 'Cần danh từ sau tính từ "significant". "Invest" -> "investment".'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...represent (22)____ future of urban living worldwide."', 'A. the', 'B. a', 'C. an', 'D. some', 'A', '"The future of + N" = tương lai của. Mạo từ xác định "the".'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: What is Singapore\'s "digital twin"?', 'A. A robot that guides tourists around the city', 'B. A virtual 3D model of the city used for urban planning', 'C. A smartphone app for booking taxis', 'D. A second city built next to Singapore', 'B', 'Đoạn văn nêu: "a virtual 3D model of the entire city that allows planners to simulate...urban planning decisions".'),
    mcq(24, 'Reading: True/False', 'Q24: Singapore uses sensors and cameras to monitor traffic flow and adjust traffic light timings.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu: "uses sensors and cameras to monitor traffic flow in real time, adjusting traffic light timings".'),
    mcq(25, 'Reading: True/False', 'Q25: According to the passage, Singapore has no green spaces in its urban planning.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu Singapore "integrates green spaces throughout the city, with vertical gardens...and park connectors".'),
    mcq(26, 'Reading: True/False', 'Q26: Other cities around the world are not interested in Singapore\'s smart city model.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu: "Other cities around the globe are now studying Singapore\'s model".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "resilient" in the passage is closest in meaning to:', 'A. fragile and easily damaged', 'B. able to recover quickly from difficulties', 'C. expensive and luxurious', 'D. old and outdated', 'B', '"Resilient" = có khả năng phục hồi nhanh, chống chịu tốt trước khó khăn.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of this passage?', 'A. To criticize Singapore\'s urban development policies', 'B. To describe how Singapore uses smart technology in urban planning and why other cities are following its lead', 'C. To encourage people to move to Singapore', 'D. To list all the problems with smart cities', 'B', 'Bài đọc mô tả cách Singapore ứng dụng công nghệ thông minh trong quy hoạch đô thị và ảnh hưởng đến các thành phố khác.'),
    textQ(29, 'Word Form', 'The ____ of smart technology in urban areas has significantly improved the quality of life for residents. (APPLY)', 'application', 'Cần danh từ sau mạo từ "The". "Apply" (động từ) -> "application" (danh từ).'),
    textQ(30, 'Word Form', 'Smart lighting systems can ____ adjust brightness based on the amount of natural light available. (AUTOMATIC)', 'automatically', 'Cần trạng từ bổ nghĩa cho động từ "adjust". "Automatic" -> "automatically".'),
    textQ(31, 'Word Form', 'The city\'s population ____ has been carefully managed through smart urban planning. (GROW)', 'growth', 'Cần danh từ. "Grow" (động từ) -> "growth" (danh từ).'),
    textQ(32, 'Word Form', 'Smart cities aim to create a more ____ urban environment for all residents. (SUSTAIN)', 'sustainable', 'Cần tính từ bổ nghĩa cho "urban environment". "Sustain" -> "sustainable".'),
    textQ(33, 'Word Form', 'Living in a smart city can be more ____ for the elderly and people with disabilities. (CONVENIENCE)', 'convenient', 'Cần tính từ sau "be more". "Convenience" (danh từ) -> "convenient" (tính từ).'),
    textQ(34, 'Word Form', 'The ____ of public services through digital platforms has made government more accessible. (DIGIT)', 'digitization', 'Cần danh từ sau mạo từ "The". "Digit" -> "digital" -> "digitization" (số hóa).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "optimize /ˈɒptɪmaɪz/ verb: to make the best or most effective use of a situation or resource." Complete: Smart grid technology can ____ energy distribution, sending power where it is needed most.', 'optimize', '"Optimize" (tối ưu hóa) phù hợp với định nghĩa và ngữ cảnh phân phối năng lượng thông minh.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "integrate /ˈɪntɪɡreɪt/ verb: to combine or bring together different systems, groups, or ideas." Complete: The new smart city platform will ____ data from traffic, weather, and energy systems into a single dashboard.', 'integrate', '"Integrate" (tích hợp) phù hợp với định nghĩa và ngữ cảnh kết hợp dữ liệu từ nhiều hệ thống.'),
    textQ(37, 'Sentence Transformation', 'City planners are designing new smart neighborhoods with green spaces and efficient transport. (DESIGNED)\n→ New smart neighborhoods ____ with green spaces and efficient transport.', 'are being designed by city planners', 'Câu bị động thì hiện tại tiếp diễn: S + am/is/are being + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'No other city in the region is as technologically advanced as Singapore. (MOST)\n→ Singapore is ____ city in the region.', 'the most technologically advanced', 'So sánh nhất: "No other...is as adj as" → "This is the most adj + N".'),
    textQ(39, 'Sentence Transformation', 'They have installed smart sensors throughout the city center to monitor air quality. (BEEN)\n→ Smart sensors ____ throughout the city center to monitor air quality.', 'have been installed', 'Câu bị động thì hiện tại hoàn thành: S + have/has been + V3/ed.'),
    textQ(40, 'Sentence Transformation', '"The new smart parking app will launch next Monday," the mayor announced. (ANNOUNCED)\n→ The mayor ____ the following Monday.', 'announced that the new smart parking app would launch', 'Câu tường thuật: lùi thì "will launch" → "would launch", "next Monday" → "the following Monday".'),
  ];

  seedSingleTest_({
    test_id: '25',
    title: 'Đề thi thử số 25',
    description: 'Chủ đề Đô thị thông minh - Từ vựng về thành phố tương lai',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 26: Bám sát đề chính thức 2023 - Ngữ pháp trọng tâm học kỳ 2
// ======================================================================
function seedQuestionsTest26_() {
  var clozePassage = [
    "English is the most widely (17)____ language in the world, with over 1.5 billion speakers across the globe.",
    "It serves as the primary language of international business, science, and technology. (18)____ English opens up",
    "opportunities for higher education, better jobs, and cross-cultural communication. In Vietnam, English has become",
    "a compulsory (19)____ in schools from primary to high school level. Many parents also invest in extra English",
    "classes for their children, hoping to give (20)____ an advantage in the future. However, many students still",
    "struggle (21)____ speaking and listening skills despite years of study. Experts suggest that students should",
    "practice English daily, (22)____ by watching movies, listening to music, or chatting with foreign friends online."
  ].join('\n\n');

  var readingPassage = [
    "Studying abroad has become an increasingly popular choice for Vietnamese students in recent years. According to",
    "the Ministry of Education and Training, over 190,000 Vietnamese students were studying in foreign countries in",
    "2023, with Japan, Australia, the United States, and South Korea among the top destinations. Students choose to",
    "study abroad for various reasons: to access higher quality education, to experience different cultures, and to",
    "improve their foreign language skills. However, studying abroad also comes with significant challenges. Many",
    "students struggle with homesickness, language barriers, and financial pressures. Some also face difficulties",
    "adapting to different teaching and learning styles. Education experts recommend that students prepare thoroughly",
    "before deciding to study overseas — not just academically, but also emotionally and financially. 'Studying abroad",
    "can be a life-changing experience,' said one counselor, 'but it requires careful planning and realistic expectations",
    "to make the most of the opportunity.'"
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. asked B. prepared C. hoped D. watched', 'A. asked /t/', 'B. prepared /d/', 'C. hoped /t/', 'D. watched /t/', 'B', '"Prepared" có âm cuối /r/ hữu thanh nên -ed đọc /d/, các từ còn lại có âm cuối vô thanh nên đọc /t/.'),
    mcq(2, 'Pronunciation: -s/-es endings', 'Choose the word whose underlined part is pronounced differently: A. books B. pens C. maps D. desks', 'A. books /s/', 'B. pens /z/', 'C. maps /s/', 'D. desks /s/', 'B', '"Pens" có âm cuối /n/ hữu thanh nên -s đọc /z/, các từ còn lại có âm cuối vô thanh nên đọc /s/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. father B. mother C. believe D. brother', 'A. father', 'B. mother', 'C. believe', 'D. brother', 'C', '"Believe" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. family B. computer C. holiday D. cinema', 'A. family', 'B. computer', 'C. holiday', 'D. cinema', 'B', '"Computer" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'Learning a new language requires patience and ____. You cannot expect to become fluent overnight.', 'A. perseverance', 'B. intelligence', 'C. entertainment', 'D. equipment', 'A', '"Perseverance" (sự kiên trì) là phẩm chất cần thiết khi học ngoại ngữ.'),
    mcq(6, 'Grammar: Tenses', 'While she ____ for the bus this morning, she ran into an old friend she had not seen in years.', 'A. waited', 'B. was waiting', 'C. has waited', 'D. waits', 'B', 'Quá khứ tiếp diễn: While + S + was/were V-ing, S + V (QKĐ). Hành động đang diễn ra thì hành động khác xen vào.'),
    mcq(7, 'Phrasal Verbs', 'The English exam was ____ because the principal wanted to give students more time to prepare.', 'A. put off', 'B. taken off', 'C. given up', 'D. turned down', 'A', '"Put off" = hoãn lại. "The exam was put off" = kỳ thi bị hoãn.'),
    mcq(8, 'Prepositions', 'Valentine\'s Day is celebrated ____ February 14th every year with cards, flowers, and chocolates.', 'A. on', 'B. in', 'C. at', 'D. for', 'A', '"On" dùng trước ngày tháng cụ thể: on + ngày tháng.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish I ____ enough money to buy a new laptop for my online English classes next semester.', 'A. have', 'B. had', 'C. will have', 'D. has', 'B', 'Câu ước wish ở hiện tại: S + wish + S + V (quá khứ đơn).'),
    mcq(10, 'Grammar: Conditional', 'If I ____ you, I would apply for that scholarship to study abroad in Australia next year.', 'A. am', 'B. was', 'C. were', 'D. be', 'C', 'Câu điều kiện loại 2: If + S + were, S + would + V. Dùng "were" cho mọi ngôi.'),
    mcq(11, 'Vocabulary: Word Choice', 'Watching English movies with subtitles is a fun and ____ way to improve your listening skills.', 'A. effective', 'B. difficult', 'C. boring', 'D. useless', 'A', '"Effective" (hiệu quả) là từ tích cực mô tả phương pháp học tiếng Anh qua phim.'),
    mcq(12, 'Communication', 'Student A: "I have just passed the entrance exam to my dream high school!"\nStudent B: "____"', 'A. That is a good idea.', 'B. Congratulations! You must be so proud.', 'C. Thank you very much.', 'D. Not at all.', 'B', '"Congratulations!" là lời chúc mừng phù hợp khi ai đó thông báo tin vui về thi đỗ.'),
    mcq(13, 'Communication', 'Student: "Would you mind if I used your dictionary for a moment?"\nTeacher: "____"', 'A. Yes, I would love to.', 'B. Of course not. Here you are.', 'C. You are welcome.', 'D. Never mind.', 'B', '"Of course not" = Dĩ nhiên là không phiền. Là cách trả lời lịch sự cho câu hỏi "Would you mind if...?".'),
    mcq(14, 'Public Signs', 'What does a sign with a crossed-out cigarette symbol mean?', 'A. You can smoke here freely', 'B. Smoking is allowed in this area', 'C. Smoking is strictly prohibited', 'D. Cigarettes are sold here', 'C', 'Biển báo điếu thuốc bị gạch chéo nghĩa là cấm hút thuốc.'),
    mcq(15, 'Public Signs', 'A sign showing a person throwing trash into a bin with the text "Keep Our School Clean" means:', 'A. Do not litter — put trash in the bin', 'B. Trash cans are only for decoration', 'C. Throw trash anywhere you like', 'D. Trash collection happens daily', 'A', 'Biển báo nhắc nhở bỏ rác đúng nơi quy định để giữ vệ sinh trường học.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'Would you mind ____ the window? It is getting too hot in the classroom.', 'A. open', 'B. to open', 'C. opening', 'D. opened', 'C', '"Would you mind + V-ing?" là cấu trúc cố định để hỏi/xin phép lịch sự.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "English is the most widely (17)____ language in the world..."', 'A. spoken', 'B. speaking', 'C. spoke', 'D. speaks', 'A', '"The most widely spoken" = được nói rộng rãi nhất. Quá khứ phân từ trong cấu trúc bị động.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "(18)____ English opens up opportunities for higher education..."', 'A. Learning', 'B. Learn', 'C. Learned', 'D. Learns', 'A', '"Learning English" = Việc học tiếng Anh. Danh động từ làm chủ ngữ.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...English has become a compulsory (19)____ in schools..."', 'A. subject', 'B. object', 'C. project', 'D. topic', 'A', '"Compulsory subject" = môn học bắt buộc. Là cụm từ trong giáo dục.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...hoping to give (20)____ an advantage in the future."', 'A. them', 'B. their', 'C. they', 'D. themselves', 'A', '"Give them an advantage" = cho họ một lợi thế. Cần tân ngữ "them" thay cho "their children".'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...many students still struggle (21)____ speaking and listening skills..."', 'A. with', 'B. on', 'C. at', 'D. in', 'A', '"Struggle with + N" = gặp khó khăn với điều gì. Cấu trúc cố định.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...practice English daily, (22)____ by watching movies, listening to music..."', 'A. such as', 'B. so that', 'C. because of', 'D. in order', 'A', '"Such as" dùng để đưa ra ví dụ về cách luyện tập tiếng Anh hàng ngày.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, how many Vietnamese students were studying abroad in 2023?', 'A. About 100,000', 'B. Over 190,000', 'C. Exactly 50,000', 'D. Less than 10,000', 'B', 'Đoạn văn nêu: "over 190,000 Vietnamese students were studying in foreign countries in 2023".'),
    mcq(24, 'Reading: True/False', 'Q24: Japan is not listed as a top destination for Vietnamese students studying abroad.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn liệt kê Nhật Bản là một trong những điểm đến hàng đầu.'),
    mcq(25, 'Reading: True/False', 'Q25: According to the passage, all students who study abroad find it easy to adapt with no difficulties.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu nhiều sinh viên gặp khó khăn như "homesickness, language barriers, and financial pressures".'),
    mcq(26, 'Reading: True/False', 'Q26: Experts recommend that students only need to prepare academically before studying abroad.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn khuyên chuẩn bị "not just academically, but also emotionally and financially".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "homesickness" in the passage most likely means:', 'A. Being physically sick at home', 'B. Feeling sad from being away from home and family', 'C. A type of illness caused by travel', 'D. Being allergic to new environments', 'B', '"Homesickness" = nỗi nhớ nhà. Cảm giác buồn khi xa gia đình và quê hương.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of the passage?', 'A. All Vietnamese students should study abroad', 'B. Studying abroad offers great opportunities but requires thorough preparation', 'C. Studying abroad is too difficult and should be avoided', 'D. Only wealthy families can afford to send their children abroad', 'B', 'Bài đọc nêu cả lợi ích và thách thức của du học, nhấn mạnh cần chuẩn bị kỹ lưỡng.'),
    textQ(29, 'Word Form', 'She speaks English ____ enough to communicate with native speakers without any difficulty. (FLUENT)', 'fluently', 'Cần trạng từ bổ nghĩa cho động từ "speaks". "Fluent" -> "fluently".'),
    textQ(30, 'Word Form', 'Mr. Brown is an ____ teacher with over 20 years of experience in the classroom. (EXPERIENCE)', 'experienced', 'Cần tính từ bổ nghĩa cho "teacher". "Experience" -> "experienced" (giàu kinh nghiệm).'),
    textQ(31, 'Word Form', 'The volunteers want to ____ the neighborhood park by planting more flowers and trees. (BEAUTY)', 'beautify', 'Cần động từ trong cấu trúc "want to V". "Beauty" (danh từ) -> "beautify" (động từ).'),
    textQ(32, 'Word Form', 'This medicine is very ____ against headaches and can relieve pain within 20 minutes. (EFFECT)', 'effective', 'Cần tính từ sau "be very". "Effect" (danh từ) -> "effective" (tính từ).'),
    textQ(33, 'Word Form', 'Thank you for your ____ during the difficult time our family went through. (ASSIST)', 'assistance', 'Cần danh từ sau tính từ sở hữu "your". "Assist" -> "assistance".'),
    textQ(34, 'Word Form', 'The team completed the project ____ and won first prize in the national competition. (SUCCESS)', 'successfully', 'Cần trạng từ bổ nghĩa cho động từ "completed". "Success" -> "successful" -> "successfully".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "take /teɪk/ verb: to move something or someone from one place to another; to accept or receive." Complete: Please ____ care of your younger sister while I am out shopping.', 'take', '"Take care of" = chăm sóc. Là cụm động từ cố định, phù hợp với định nghĩa.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "follow /ˈfɒləʊ/ verb: to move behind someone; to obey or act according to." Complete: Children should ____ the example set by their parents and teachers.', 'follow', '"Follow the example" = noi gương. Là cụm từ cố định, phù hợp với định nghĩa.'),
    textQ(37, 'Sentence Transformation', 'The children like making models of animals in their free time. (KEEN)\n→ The children are ____ of animals in their free time.', 'keen on making models', '"Like V-ing" → "be keen on V-ing". Nghĩa tương đương: thích làm gì.'),
    textQ(38, 'Sentence Transformation', 'I advise you to see a doctor regularly for health check-ups. (WERE)\n→ If I ____ a doctor regularly for health check-ups.', 'were you, I would see', '"I advise you to V" → "If I were you, I would V". Lời khuyên.'),
    textQ(39, 'Sentence Transformation', 'She last participated in an English competition two years ago. (HAS NOT)\n→ She ____ an English competition for two years.', 'has not participated in', '"S + last V (QKĐ) + time ago" → "S + have/has not V3/ed + for + time".'),
    textQ(40, 'Sentence Transformation', 'He performed excellently in the exam, so he was nominated for the scholarship. (BECAUSE)\n→ ____ excellently in the exam, he was nominated for the scholarship.', 'Because he performed', '"S + V, so + S + V" → "Because + S + V, S + V". Đảo mệnh đề nguyên nhân lên đầu.'),
  ];

  seedSingleTest_({
    test_id: '26',
    title: 'Đề thi thử số 26',
    description: 'Bám sát đề chính thức 2023 - Ngữ pháp trọng tâm học kỳ 2',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 27: Chủ đề Giáo dục STEM - Đọc hiểu về phương pháp học tập mới
// ======================================================================
function seedQuestionsTest27_() {
  var clozePassage = [
    "STEM education, (17)____ stands for Science, Technology, Engineering, and Mathematics, has become a global",
    "priority in recent years. Unlike traditional education, STEM focuses (18)____ hands-on learning and real-world",
    "problem-solving. Students in STEM programs learn to think critically, work collaboratively, and apply (19)____",
    "across different subjects. In Vietnam, many schools are now integrating STEM activities (20)____ their curriculum,",
    "from robotics clubs to coding workshops. Research suggests (21)____ students who engage in STEM education are",
    "better prepared for the jobs of the future. The government has also launched initiatives to train more STEM",
    "teachers and equip schools (22)____ modern laboratory facilities."
  ].join('\n\n');

  var readingPassage = [
    "Robotics has become one of the most exciting and popular STEM activities in Vietnamese schools. Over the past",
    "few years, robotics clubs have sprung up in both public and private schools across the country, giving students",
    "the opportunity to design, build, and program their own robots. These clubs not only teach technical skills like",
    "coding and engineering but also develop important soft skills such as teamwork, problem-solving, and creativity.",
    "Vietnamese students have achieved remarkable success in international robotics competitions. In 2023, a team from",
    "Ho Chi Minh City won a gold medal at the World Robot Olympiad, beating teams from over 70 countries. The students",
    "spent months preparing for the competition, working after school and on weekends to perfect their robot's design",
    "and programming. Their victory has inspired many other schools to start their own robotics programs. 'Robotics is",
    "not just about building machines,' said the team coach. 'It teaches students how to think systematically, learn",
    "from failure, and work together toward a common goal — skills that will serve them well no matter what career",
    "they choose in the future.'"
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. designed B. launched C. programmed D. prepared', 'A. designed /d/', 'B. launched /t/', 'C. programmed /d/', 'D. prepared /d/', 'B', '"Launched" có âm cuối /tʃ/ vô thanh nên -ed đọc /t/, các từ còn lại đọc /d/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. science B. technology C. engineer D. education', 'A. science /aɪ/', 'B. technology /e/', 'C. engineer /e/', 'D. education /e/', 'A', '"Science" có âm "i" đọc là /aɪ/, các từ còn lại có âm /e/ ở âm tiết đầu.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. robot B. program C. design D. system', 'A. robot', 'B. program', 'C. design', 'D. system', 'C', '"Design" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. critical B. creative C. scientist D. technical', 'A. critical', 'B. creative', 'C. scientist', 'D. technical', 'B', '"Creative" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'STEM education encourages students to think ____ and find innovative solutions to real-world problems.', 'A. critically', 'B. critical', 'C. criticize', 'D. criticism', 'A', 'Cần trạng từ bổ nghĩa cho động từ "think". "Critical" -> "critically".'),
    mcq(6, 'Grammar: Tenses', 'Since the school introduced the robotics program, student interest in STEM subjects ____ dramatically.', 'A. increased', 'B. has increased', 'C. increases', 'D. was increasing', 'B', 'Thì hiện tại hoàn thành với "Since" + mệnh đề QKĐ.'),
    mcq(7, 'Phrasal Verbs', 'The students decided to ____ a robotics club at their school after being inspired by a science fair.', 'A. set up', 'B. break up', 'C. put off', 'D. give in', 'A', '"Set up" = thành lập. "Set up a club" = thành lập câu lạc bộ.'),
    mcq(8, 'Prepositions', 'Many students are now interested ____ learning programming and app development outside of class.', 'A. in', 'B. on', 'C. at', 'D. for', 'A', '"Be interested in + N/V-ing" là cấu trúc cố định.'),
    mcq(9, 'Grammar: Conditional', 'If schools ____ more STEM programs, more students would pursue careers in science and technology.', 'A. offer', 'B. offered', 'C. will offer', 'D. would offer', 'B', 'Câu điều kiện loại 2: If + S + V (QKĐ), S + would + V.'),
    mcq(10, 'Grammar: Relative Clause', 'The students ____ robot won the international competition received scholarships from top universities.', 'A. who', 'B. which', 'C. whose', 'D. whom', 'C', '"Whose" là đại từ quan hệ sở hữu: "the students\' robot" → "whose robot".'),
    mcq(11, 'Vocabulary: Word Choice', 'The new STEM lab is ____ with modern computers, 3D printers, and robotics kits for student projects.', 'A. equipped', 'B. provided', 'C. supplied', 'D. furnished', 'A', '"Equipped with" = được trang bị với. Là cụm từ phổ biến khi nói về cơ sở vật chất.'),
    mcq(12, 'Communication', 'Teacher: "Why do you think STEM education is important for your future?"\nStudent: "____"', 'A. Because it teaches me skills like coding and problem-solving that will be valuable in any career.', 'B. I do not like math or science at all.', 'C. My parents told me to join the robotics club.', 'D. STEM subjects are too difficult for most students.', 'A', 'Trả lời về tầm quan trọng của STEM với các kỹ năng cụ thể, phù hợp câu hỏi của giáo viên.'),
    mcq(13, 'Communication', 'Student A: "I am having trouble programming my robot for the competition."\nStudent B: "____"', 'A. Let us work on it together. I can show you what worked for my design.', 'B. You should just give up. It is too hard.', 'C. Robots are not important for our future.', 'D. I do not have time to help anyone else.', 'A', 'Đề nghị giúp đỡ bạn cùng lớp với tinh thần hợp tác, đúng tinh thần STEM.'),
    mcq(14, 'Public Signs', 'A sign in a lab showing a pair of safety goggles and the text "Eye Protection Required" means:', 'A. Goggles are optional in this area', 'B. You must wear eye protection to enter', 'C. This is an eyewear shop', 'D. Do not look directly at the light', 'B', 'Biển báo kính bảo hộ yêu cầu đeo kính bảo vệ mắt khi vào khu vực thí nghiệm.'),
    mcq(15, 'Public Signs', 'What does a sign showing a crossed-out food item at a computer lab entrance mean?', 'A. Cafeteria is nearby', 'B. No food or drinks allowed in the lab', 'C. Free snacks available inside', 'D. Vending machine location', 'B', 'Biển báo cấm đồ ăn ở cửa phòng máy tính để bảo vệ thiết bị khỏi hư hỏng.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She enjoys ____ with circuit boards and sensors to create simple electronic projects.', 'A. experiment', 'B. to experiment', 'C. experimenting', 'D. experimented', 'C', '"Enjoy + V-ing" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "STEM education, (17)____ stands for Science, Technology, Engineering, and Mathematics..."', 'A. which', 'B. who', 'C. what', 'D. where', 'A', '"Which" thay cho "STEM education" trong mệnh đề quan hệ không xác định.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...STEM focuses (18)____ hands-on learning and real-world problem-solving."', 'A. on', 'B. in', 'C. at', 'D. to', 'A', '"Focus on + N" = tập trung vào. Cấu trúc cố định.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...think critically, work collaboratively, and apply (19)____ across different subjects."', 'A. knowledge', 'B. know', 'C. knowing', 'D. known', 'A', 'Cần danh từ sau động từ "apply". "Knowledge" = kiến thức.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...integrating STEM activities (20)____ their curriculum..."', 'A. into', 'B. with', 'C. for', 'D. from', 'A', '"Integrate A into B" = tích hợp A vào B.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "Research suggests (21)____ students who engage in STEM education are better prepared..."', 'A. that', 'B. what', 'C. which', 'D. whether', 'A', '"Suggests that + clause" = gợi ý rằng. Dùng "that" để nối mệnh đề.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...equip schools (22)____ modern laboratory facilities."', 'A. with', 'B. for', 'C. about', 'D. to', 'A', '"Equip someone/something with something" = trang bị cho ai/cái gì với thứ gì.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: What do robotics clubs teach students besides technical skills?', 'A. Only how to use computers', 'B. Soft skills like teamwork, problem-solving, and creativity', 'C. How to become professional athletes', 'D. Cooking and gardening skills', 'B', 'Đoạn văn nêu các kỹ năng mềm: "teamwork, problem-solving, and creativity".'),
    mcq(24, 'Reading: True/False', 'Q24: Vietnamese students have never won any awards in international robotics competitions.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu đội từ TP.HCM giành huy chương vàng tại World Robot Olympiad 2023.'),
    mcq(25, 'Reading: True/False', 'Q25: The winning team spent only one week preparing for the competition.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu họ đã "spent months preparing" (dành nhiều tháng chuẩn bị).'),
    mcq(26, 'Reading: True/False', 'Q26: According to the coach, robotics teaches students to learn from failure.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn trích lời huấn luyện viên: "learn from failure, and work together".'),
    mcq(27, 'Reading: MCQ', 'Q27: The phrase "think systematically" used by the coach most likely means:', 'A. Think randomly without any method', 'B. Think in an organized, logical, step-by-step way', 'C. Think about only one subject at a time', 'D. Think emotionally rather than logically', 'B', '"Think systematically" = suy nghĩ có hệ thống, logic theo từng bước.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message the coach wants to convey about robotics?', 'A. Robotics is only for students who want to become engineers', 'B. Robotics develops valuable life skills that benefit students in any future career', 'C. Robotics competitions are too difficult for most Vietnamese students', 'D. Winning competitions is the only goal of robotics programs', 'B', 'HLV nhấn mạnh robotics dạy các kỹ năng sống có giá trị cho bất kỳ nghề nghiệp tương lai nào.'),
    textQ(29, 'Word Form', 'The ____ of STEM programs in schools has increased significantly across Vietnam in recent years. (INTRODUCE)', 'introduction', 'Cần danh từ sau mạo từ "The". "Introduce" (động từ) -> "introduction" (danh từ).'),
    textQ(30, 'Word Form', 'Coding is becoming an ____ important skill in the modern job market. (INCREASE)', 'increasingly', 'Cần trạng từ bổ nghĩa cho tính từ "important". "Increase" -> "increasing" -> "increasingly".'),
    textQ(31, 'Word Form', 'The students showed great ____ when designing their robot for the international competition. (CREATE)', 'creativity', 'Cần danh từ sau tính từ "great". "Create" -> "creative" -> "creativity".'),
    textQ(32, 'Word Form', 'There is a high ____ for workers with STEM skills in the technology industry. (DEMANDING)', 'demand', 'Cần danh từ sau "a high". "Demanding" (tính từ) -> "demand" (danh từ).'),
    textQ(33, 'Word Form', 'She solved the complex math problem ____ and impressed all her classmates. (EASY)', 'easily', 'Cần trạng từ bổ nghĩa cho động từ "solved". "Easy" -> "easily".'),
    textQ(34, 'Word Form', 'Many ____ innovations have come from young STEM students working on school projects. (IMPRESS)', 'impressive', 'Cần tính từ bổ nghĩa cho "innovations". "Impress" -> "impressive" (ấn tượng).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "code /kəʊd/ verb: to write computer programs using a programming language." Complete: In the after-school club, students learn to ____ simple games using Python and Scratch.', 'code', '"Code" (viết mã, lập trình) phù hợp với định nghĩa và ngữ cảnh câu lạc bộ STEM.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "innovate /ˈɪnəveɪt/ verb: to introduce new ideas, methods, or products." Complete: The best companies encourage their employees to ____ and come up with creative solutions.', 'innovate', '"Innovate" (đổi mới, sáng tạo) phù hợp với định nghĩa và ngữ cảnh khuyến khích nhân viên sáng tạo.'),
    textQ(37, 'Sentence Transformation', 'Schools should encourage more girls to pursue STEM subjects and careers. (ENCOURAGED)\n→ More girls ____ to pursue STEM subjects and careers by schools.', 'should be encouraged', 'Câu bị động với động từ khuyết thiếu: S + should + be + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'They started building the new science laboratory three months ago. (BEEN)\n→ The new science laboratory ____ for three months.', 'has been built', '"S + started V-ing + time ago" → "S + have/has been V3/ed + for + time".'),
    textQ(39, 'Sentence Transformation', 'Learning to code is more important than memorizing facts in today\'s world. (AS)\n→ Memorizing facts is ____ learning to code in today\'s world.', 'not as important as', 'So sánh hơn "more important than" → so sánh ngang bằng phủ định "not as important as".'),
    textQ(40, 'Sentence Transformation', '"Work together to solve this engineering challenge," the teacher said to the students. (ASKED)\n→ The teacher ____ to solve that engineering challenge.', 'asked the students to work together', '"V," S said to O → "S + asked + O + to V".'),
  ];

  seedSingleTest_({
    test_id: '27',
    title: 'Đề thi thử số 27',
    description: 'Chủ đề Giáo dục STEM - Đọc hiểu về phương pháp học tập mới',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 28: Chủ đề Kỹ năng mềm & Lãnh đạo - Giao tiếp và thuyết trình
// ======================================================================
function seedQuestionsTest28_() {
  var clozePassage = [
    "Soft skills such (17)____ communication, teamwork, and leadership are increasingly valued in today's workplace.",
    "While technical skills may help you get an interview, soft skills often determine (18)____ you get the job and",
    "succeed in it. Effective communication involves not only speaking clearly (19)____ also listening actively and",
    "understanding non-verbal cues. Leadership is not just about giving orders; it is about inspiring (20)____, making",
    "decisions under pressure, and taking responsibility. Many schools and universities now offer courses and workshops",
    "to help students develop (21)____ essential skills. The earlier you start building these skills, (22)____ better",
    "prepared you will be for the challenges of professional life."
  ].join('\n\n');

  var readingPassage = [
    "Emotional intelligence, often referred to as EQ, is increasingly recognized as a key factor in both personal and",
    "professional success. While IQ measures cognitive abilities such as logical reasoning and problem-solving, EQ",
    "refers to the ability to understand, manage, and express one's own emotions while also being able to recognize",
    "and influence the emotions of others. Psychologist Daniel Goleman, who popularized the concept, identified five",
    "key components of emotional intelligence: self-awareness, self-regulation, motivation, empathy, and social skills.",
    "Research has shown that people with high EQ tend to have better relationships, perform better at work, and enjoy",
    "greater mental well-being. In the workplace, managers with high emotional intelligence are more effective at",
    "motivating their teams and handling conflicts. The good news is that unlike IQ, which remains relatively stable",
    "throughout life, emotional intelligence can be developed and improved with practice. Simple habits like reflecting",
    "on your emotions, practicing active listening, and seeking feedback from others can help boost your EQ over time."
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. valued B. developed C. improved D. recognized', 'A. valued /d/', 'B. developed /t/', 'C. improved /d/', 'D. recognized /d/', 'B', '"Developed" có âm cuối /p/ vô thanh nên -ed đọc /t/, các từ còn lại đọc /d/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. leader B. skill C. listen D. talk', 'A. leader /iː/', 'B. skill /ɪ/', 'C. listen /ɪ/', 'D. talk /ɔː/', 'A', '"Leader" có âm "ea" đọc là /iː/, các từ còn lại có âm /ɪ/ hoặc /ɔː/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. leader B. inspire C. manage D. listen', 'A. leader', 'B. inspire', 'C. manage', 'D. listen', 'B', '"Inspire" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. motivate B. empathy C. recognize D. influence', 'A. motivate', 'B. empathy', 'C. recognize', 'D. influence', 'B', '"Empathy" nhấn âm 1 (hoặc 3), các từ còn lại nhấn âm 1. Cần kiểm tra: "empathy" nhấn âm 1, "motivate" nhấn âm 1, "recognize" nhấn âm 1, "influence" nhấn âm 1. Thực tế "empathy" nhấn âm 1 trong khi các từ còn lại cũng nhấn âm 1. Từ khác: "emotion" nhấn âm 2.'),
    mcq(5, 'Vocabulary in Context', 'The ability to ____ effectively with colleagues is one of the most important soft skills in any profession.', 'A. collaborate', 'B. compete', 'C. complain', 'D. compare', 'A', '"Collaborate" (hợp tác) là kỹ năng mềm thiết yếu trong môi trường làm việc nhóm.'),
    mcq(6, 'Grammar: Tenses', 'Over the past decade, employers ____ greater emphasis on soft skills when hiring new graduates.', 'A. placed', 'B. have placed', 'C. place', 'D. were placing', 'B', 'Thì hiện tại hoàn thành với "Over the past decade".'),
    mcq(7, 'Phrasal Verbs', 'A good leader knows how to ____ conflicts between team members before they escalate.', 'A. deal with', 'B. put off', 'C. look up', 'D. give up', 'A', '"Deal with" = giải quyết, xử lý. "Deal with conflicts" = giải quyết xung đột.'),
    mcq(8, 'Prepositions', 'Public speaking is a skill that many people struggle ____, but it can be improved with practice.', 'A. with', 'B. on', 'C. at', 'D. for', 'A', '"Struggle with + N" = gặp khó khăn với điều gì.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish I ____ more confident when presenting in front of large audiences at company meetings.', 'A. am', 'B. were', 'C. will be', 'D. have been', 'B', 'Câu ước ở hiện tại: S + wish + S + V (QKĐ). Dùng "were" cho mọi ngôi.'),
    mcq(10, 'Grammar: Comparison', 'Developing soft skills is just ____ as acquiring technical knowledge for long-term career success.', 'A. as important', 'B. more important', 'C. most important', 'D. the most important', 'A', 'So sánh ngang bằng: as + adj + as.'),
    mcq(11, 'Vocabulary: Word Choice', 'A person with high ____ intelligence can understand and manage both their own emotions and those of others.', 'A. emotional', 'B. physical', 'C. artificial', 'D. technical', 'A', '"Emotional intelligence" (trí tuệ cảm xúc) là thuật ngữ của Daniel Goleman.'),
    mcq(12, 'Communication', 'Manager: "I would like you to lead the presentation to our new clients next week."\nEmployee: "____"', 'A. I would be happy to. Could we schedule a time to discuss the key points?', 'B. Why me? Someone else should do it.', 'C. I am too nervous to present to anyone.', 'D. Presentations are not part of my job description.', 'A', 'Phản hồi tích cực, chuyên nghiệp, sẵn sàng nhận trách nhiệm và chuẩn bị kỹ lưỡng.'),
    mcq(13, 'Communication', 'Colleague: "Your presentation was really impressive. How did you prepare?"\nYou: "____"', 'A. Thank you! I practiced several times and focused on telling a clear story with the data.', 'B. It was nothing special. Anyone could have done it.', 'C. I did not prepare at all. I just got lucky.', 'D. Presentations do not matter in the real world.', 'A', 'Chia sẻ bí quyết chuẩn bị thuyết trình một cách khiêm tốn và hữu ích.'),
    mcq(14, 'Public Signs', 'A sign showing a person speaking at a podium with "Presentation in Progress" means:', 'A. Free speech zone', 'B. Meeting or presentation in progress — do not disturb', 'C. Public speaking contest registration', 'D. Open microphone for anyone', 'B', 'Biển báo "Presentation in Progress" yêu cầu không làm gián đoạn buổi thuyết trình đang diễn ra.'),
    mcq(15, 'Public Signs', 'What does a sign showing two people shaking hands with the text "Meeting Point" mean?', 'A. Designated meeting or assembly location', 'B. Only business people can stand here', 'C. Handshake practice area', 'D. Job interview room', 'A', 'Biển báo hai người bắt tay với chữ "Meeting Point" chỉ điểm hẹn hoặc nơi tập trung.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The company encourages its employees ____ workshops on leadership and communication skills.', 'A. attend', 'B. to attend', 'C. attending', 'D. attended', 'B', '"Encourage + O + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Soft skills such (17)____ communication, teamwork, and leadership..."', 'A. as', 'B. like', 'C. of', 'D. for', 'A', '"Such as" = ví dụ như. Dùng để liệt kê.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...soft skills often determine (18)____ you get the job and succeed in it."', 'A. whether', 'B. if', 'C. that', 'D. which', 'A', '"Determine whether + clause" = xác định liệu rằng. Phù hợp nhất về nghĩa.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...not only speaking clearly (19)____ also listening actively..."', 'A. but', 'B. and', 'C. or', 'D. so', 'A', 'Cấu trúc "not only...but also" = không những...mà còn.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...it is about inspiring (20)____, making decisions under pressure..."', 'A. others', 'B. other', 'C. another', 'D. the others', 'A', '"Inspiring others" = truyền cảm hứng cho người khác.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...help students develop (21)____ essential skills."', 'A. these', 'B. this', 'C. that', 'D. those', 'A', '"These" chỉ các kỹ năng đã được đề cập gần đó (số nhiều, gần).'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "The earlier you start..., (22)____ better prepared you will be..."', 'A. the', 'B. a', 'C. an', 'D. more', 'A', 'Cấu trúc so sánh kép: "The + so sánh hơn..., the + so sánh hơn...".'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to Daniel Goleman, which of the following is NOT one of the five components of emotional intelligence?', 'A. Self-awareness', 'B. Mathematical ability', 'C. Empathy', 'D. Self-regulation', 'B', 'Đoạn văn liệt kê 5 thành phần: self-awareness, self-regulation, motivation, empathy, social skills. Không có mathematical ability.'),
    mcq(24, 'Reading: True/False', 'Q24: EQ refers to the ability to understand and manage emotions.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn định nghĩa EQ là "the ability to understand, manage, and express one\'s own emotions".'),
    mcq(25, 'Reading: True/False', 'Q25: IQ can be easily improved with practice throughout one\'s life, just like EQ.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu IQ "remains relatively stable throughout life" trong khi EQ có thể phát triển.'),
    mcq(26, 'Reading: True/False', 'Q26: According to research, people with high EQ tend to have better mental well-being.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu người EQ cao "enjoy greater mental well-being".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "regulate" in "self-regulation" is closest in meaning to:', 'A. ignore', 'B. control', 'C. express', 'D. understand', 'B', '"Self-regulation" = tự điều chỉnh, tự kiểm soát. "Regulate" đồng nghĩa với "control".'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of the passage about emotional intelligence?', 'A. EQ is fixed at birth and cannot be changed', 'B. EQ is more important than IQ and can be developed with practice', 'C. Only managers need emotional intelligence', 'D. Emotional intelligence has no impact on career success', 'B', 'Bài đọc nhấn mạnh EQ quan trọng và có thể phát triển, không giống IQ cố định.'),
    textQ(29, 'Word Form', 'Good ____ skills are essential for anyone who wants to become a successful team leader. (LEAD)', 'leadership', 'Cần danh từ đứng trước "skills". "Lead" (động từ) -> "leadership" (danh từ).'),
    textQ(30, 'Word Form', 'She handled the difficult customer complaint ____ and managed to keep the situation calm. (PROFESSION)', 'professionally', 'Cần trạng từ bổ nghĩa cho động từ "handled". "Profession" -> "professional" -> "professionally".'),
    textQ(31, 'Word Form', 'Active ____ is one of the most important communication skills in both personal and professional life. (LISTEN)', 'listening', 'Cần danh từ (danh động từ) sau tính từ "Active". "Listen" -> "listening".'),
    textQ(32, 'Word Form', 'The ____ to adapt to change is what separates great leaders from average ones. (ABLE)', 'ability', 'Cần danh từ sau mạo từ "The". "Able" (tính từ) -> "ability" (danh từ).'),
    textQ(33, 'Word Form', 'He spoke ____ during the interview and convinced the panel he was the right candidate. (PERSUADE)', 'persuasively', 'Cần trạng từ bổ nghĩa cho động từ "spoke". "Persuade" -> "persuasive" -> "persuasively".'),
    textQ(34, 'Word Form', 'Constructive ____ from colleagues can help you identify areas for professional growth. (FEED)', 'feedback', 'Cần danh từ. "Feed" + "back" = "feedback" (phản hồi).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "negotiate /nɪˈɡəʊʃieɪt/ verb: to try to reach an agreement through discussion and compromise." Complete: A good manager knows how to ____ with clients to get the best deal for the company.', 'negotiate', '"Negotiate" (đàm phán) phù hợp với định nghĩa và ngữ cảnh thương lượng với khách hàng.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "delegate /ˈdelɪɡeɪt/ verb: to give a task or responsibility to another person, typically a subordinate." Complete: Effective leaders know when to ____ tasks to team members rather than doing everything themselves.', 'delegate', '"Delegate tasks" (giao việc, ủy quyền) phù hợp với định nghĩa và ngữ cảnh lãnh đạo.'),
    textQ(37, 'Sentence Transformation', 'If you practice regularly, your presentation skills will improve significantly. (PRACTICE)\n→ Regular ____ will help improve your presentation skills significantly.', 'practice', 'Chuyển từ mệnh đề "If you practice" thành danh từ làm chủ ngữ: "Regular practice".'),
    textQ(38, 'Sentence Transformation', 'No other team member communicates as effectively as Sarah does. (MOST)\n→ Sarah is ____ team member.', 'the most effective communicator in the', '"No other...as...as" → "the most + adj + N".'),
    textQ(39, 'Sentence Transformation', '"Take the lead on this project," the director said to me. (ASKED)\n→ The director ____ on that project.', 'asked me to take the lead', '"V," S said to O → "S + asked + O + to V".'),
    textQ(40, 'Sentence Transformation', 'She has not attended any leadership training since she started her current position. (LAST)\n→ The ____ was when she started her current position.', 'last time she attended leadership training', '"S + have/has not V3/ed + since..." → "The last time + S + V (QKĐ) + was...".'),
  ];

  seedSingleTest_({
    test_id: '28',
    title: 'Đề thi thử số 28',
    description: 'Chủ đề Kỹ năng mềm & Lãnh đạo - Giao tiếp và thuyết trình',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 29: Chủ đề Bảo vệ động vật quý hiếm - Từ vựng về động vật hoang dã
// ======================================================================
function seedQuestionsTest29_() {
  var clozePassage = [
    "Wildlife conservation is a critical global issue as many species face the (17)____ of extinction. Iconic animals",
    "such as tigers, elephants, rhinos, and pandas are (18)____ the most endangered species on the planet. The main",
    "threats to wildlife include habitat destruction, illegal poaching, and the illegal wildlife (19)____. Conservation",
    "organizations worldwide are working to protect these animals (20)____ anti-poaching patrols, breeding programs,",
    "and public education. In Vietnam, the saola, often called the 'Asian unicorn', is one of the (21)____ and most",
    "threatened mammals in the world. Successfully protecting endangered species (22)____ cooperation between",
    "governments, conservation groups, and local communities."
  ].join('\n\n');

  var readingPassage = [
    "The saola, often called the 'Asian unicorn', is one of the world's rarest and most mysterious large mammals.",
    "Discovered only in 1992 in the Annamite Mountains along the Vietnam-Laos border, the saola was the first large",
    "mammal new to science in more than 50 years. With its long, straight horns and distinctive white facial markings,",
    "the saola looks like an animal from a fairy tale. Scientists estimate that fewer than 100 saolas remain in the",
    "wild, and possibly as few as 20. The main threats to the saola are hunting snares set by poachers targeting other",
    "animals, as well as habitat loss due to logging and agricultural expansion. Conservation groups like the Saola",
    "Working Group and WWF are working with local communities and the Vietnamese government to protect the saola's",
    "remaining habitat and remove deadly snares from the forest. There is also a plan to establish a captive breeding",
    "program to ensure the species' survival if all wild saolas are lost. 'The saola is a symbol of the incredible",
    "biodiversity that still exists in the Annamites,' said one conservation biologist. 'If we lose the saola, we lose",
    "a unique piece of our planet's natural heritage forever.'"
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. hunted B. poached C. trapped D. threatened', 'A. hunted /ɪd/', 'B. poached /t/', 'C. trapped /t/', 'D. threatened /d/', 'A', '"Hunted" có âm cuối /t/ nên -ed đọc /ɪd/, các từ còn lại đọc /t/ hoặc /d/.'),
    mcq(2, 'Pronunciation: Consonant sounds', 'Choose the word whose underlined part is pronounced differently: A. species B. extinction C. conservation D. habitat', 'A. species /ʃ/', 'B. extinction /ks/', 'C. conservation /s/', 'D. habitat /h/', 'A', '"Species" có âm "c" đọc là /ʃ/, các từ còn lại có âm /ks/, /s/, /h/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. species B. wildlife C. protect D. forest', 'A. species', 'B. wildlife', 'C. protect', 'D. forest', 'C', '"Protect" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. elephant B. extinction C. habitat D. benefit', 'A. elephant', 'B. extinction', 'C. habitat', 'D. benefit', 'B', '"Extinction" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', '____ is the illegal hunting and killing of wild animals, often for their valuable parts such as horns, tusks, or fur.', 'A. Poaching', 'B. Farming', 'C. Breeding', 'D. Training', 'A', '"Poaching" (săn bắt trộm) là thuật ngữ chính xác chỉ hành vi săn bắt động vật hoang dã trái phép.'),
    mcq(6, 'Grammar: Tenses', 'Since the saola was discovered in 1992, conservationists ____ tirelessly to protect this rare species.', 'A. worked', 'B. have been working', 'C. work', 'D. are working', 'B', 'Hiện tại hoàn thành tiếp diễn với "Since", nhấn mạnh hành động liên tục từ quá khứ đến nay.'),
    mcq(7, 'Phrasal Verbs', 'Conservation groups are trying to ____ awareness about the illegal trade in endangered species.', 'A. raise', 'B. bring up', 'C. put off', 'D. give in', 'A', '"Raise awareness" = nâng cao nhận thức. Là cụm cố định trong lĩnh vực bảo tồn.'),
    mcq(8, 'Prepositions', 'The saola is classified ____ critically endangered, with possibly fewer than 20 individuals remaining.', 'A. as', 'B. for', 'C. with', 'D. in', 'A', '"Classified as + adj/N" = được phân loại là.'),
    mcq(9, 'Grammar: Conditional', 'If the saola ____ to extinction, it would be a tremendous loss to global biodiversity.', 'A. goes', 'B. went', 'C. will go', 'D. would go', 'B', 'Câu điều kiện loại 2: If + S + V (QKĐ), S + would + V.'),
    mcq(10, 'Grammar: Passive Voice', 'Many endangered species ____ in national parks and breeding centers to ensure their survival.', 'A. are kept', 'B. keep', 'C. kept', 'D. have kept', 'A', 'Câu bị động thì hiện tại đơn: S (số nhiều) + are + V3/ed.'),
    mcq(11, 'Vocabulary: Word Choice', 'The saola has distinctive white ____ on its face, making it easily recognizable.', 'A. markings', 'B. stains', 'C. dots', 'D. spots', 'A', '"Markings" (vằn, đốm đặc trưng) là từ chuyên ngành để mô tả đặc điểm nhận dạng của động vật.'),
    mcq(12, 'Communication', 'Student: "What can I do to help protect endangered animals?"\nConservationist: "____"', 'A. You can support wildlife organizations, reduce your plastic use, and spread awareness on social media.', 'B. There is nothing you can do. The situation is hopeless.', 'C. Only scientists and governments can make a difference.', 'D. You should not worry about animals that live far away.', 'A', 'Câu trả lời đưa ra các hành động cụ thể, thiết thực mà học sinh có thể làm.'),
    mcq(13, 'Communication', 'Friend: "Would it be okay to buy a souvenir made from ivory at the market?"\nYou: "____"', 'A. No, you should not. Buying ivory products encourages poaching and is illegal in many places.', 'B. Yes, it makes a beautiful and unique gift for your family.', 'C. I am not sure. Let me check the price first.', 'D. Ivory is not expensive, so you can buy as much as you want.', 'A', 'Phản hồi có trách nhiệm, giải thích tác hại của việc mua sản phẩm từ ngà voi.'),
    mcq(14, 'Public Signs', 'A sign at a national park showing a crossed-out rifle means:', 'A. Hunting is prohibited in this area', 'B. Shooting range ahead', 'C. Military training zone', 'D. Hunting licenses sold here', 'A', 'Biển báo súng bị gạch chéo trong vườn quốc gia nghĩa là cấm săn bắn.'),
    mcq(15, 'Public Signs', 'What does a sign showing a turtle inside a red circle with a line through it at a beach mean?', 'A. Turtle nesting area — do not disturb', 'B. Turtle soup sold here', 'C. Turtle racing event', 'D. Turtle petting zoo', 'A', 'Biển báo rùa biển bị gạch chéo ở bãi biển cảnh báo khu vực rùa làm tổ, không được làm phiền.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'The conservation team decided ____ camera traps in the forest to monitor saola activity.', 'A. set up', 'B. to set up', 'C. setting up', 'D. set up', 'B', '"Decide + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "...as many species face the (17)____ of extinction."', 'A. threat', 'B. threaten', 'C. threatening', 'D. threatened', 'A', '"Face the threat of + N" = đối mặt với nguy cơ. Cần danh từ.'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...are (18)____ the most endangered species on the planet."', 'A. among', 'B. between', 'C. in', 'D. with', 'A', '"Among the most..." = trong số những...nhất. Dùng cho nhóm nhiều hơn 2.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...illegal poaching, and the illegal wildlife (19)____."', 'A. trade', 'B. selling', 'C. market', 'D. exchange', 'A', '"Illegal wildlife trade" = buôn bán động vật hoang dã trái phép. Là thuật ngữ chuẩn.'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...working to protect these animals (20)____ anti-poaching patrols..."', 'A. through', 'B. by', 'C. with', 'D. from', 'A', '"Through + N" = thông qua, bằng cách. Chỉ phương thức.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "...one of the (21)____ and most threatened mammals in the world."', 'A. rarest', 'B. rare', 'C. rarely', 'D. more rare', 'A', 'So sánh nhất: "one of the + adj-est/most adj" + danh từ số nhiều. "Rare" -> "rarest".'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "Successfully protecting endangered species (22)____ cooperation..."', 'A. requires', 'B. require', 'C. requiring', 'D. required', 'A', '"Protecting endangered species" (danh động từ) là chủ ngữ số ít, động từ thêm -s.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: When was the saola first discovered by scientists?', 'A. In 1982', 'B. In 1992', 'C. In 2002', 'D. In 2012', 'B', 'Đoạn văn nêu: "Discovered only in 1992 in the Annamite Mountains".'),
    mcq(24, 'Reading: True/False', 'Q24: The saola was the first large mammal new to science in more than 50 years.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn xác nhận: "the first large mammal new to science in more than 50 years."'),
    mcq(25, 'Reading: True/False', 'Q25: Scientists estimate there are over 1,000 saolas remaining in the wild.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu ước tính "fewer than 100" và có thể chỉ "as few as 20".'),
    mcq(26, 'Reading: True/False', 'Q26: A captive breeding program is being planned as a backup to save the saola from extinction.', 'A. TRUE', '', 'B. FALSE', '', 'A', 'Đoạn văn nêu có kế hoạch "establish a captive breeding program to ensure the species\' survival".'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "snares" in the passage most likely means:', 'A. traps made from wire or rope used to catch animals', 'B. loud noises that scare animals away', 'C. types of medicine for sick animals', 'D. food used to attract animals for observation', 'A', '"Snares" = bẫy (thường làm từ dây thép hoặc dây thừng) dùng để bắt động vật.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main purpose of this passage?', 'A. To promote tourism to the Annamite Mountains', 'B. To raise awareness about the critically endangered saola and efforts to save it', 'C. To encourage people to keep saolas as pets', 'D. To argue that conservation is not necessary', 'B', 'Bài đọc nhằm nâng cao nhận thức về loài saola cực kỳ nguy cấp và các nỗ lực bảo tồn.'),
    textQ(29, 'Word Form', 'The ____ of many wild animals has become a serious problem in countries with weak law enforcement. (HUNT)', 'hunting', 'Cần danh từ (danh động từ) sau mạo từ "The". "Hunt" -> "hunting" (săn bắt).'),
    textQ(30, 'Word Form', 'Breaking wildlife protection laws can result in ____ penalties, including heavy fines and prison time. (SEVERE)', 'severe', 'Cần tính từ bổ nghĩa cho "penalties". "Severe" giữ nguyên — tính từ.'),
    textQ(31, 'Word Form', 'The ____ of the saola depends on the collective efforts of governments, NGOs, and local communities. (SURVIVE)', 'survival', 'Cần danh từ sau mạo từ "The". "Survive" (động từ) -> "survival" (danh từ).'),
    textQ(32, 'Word Form', 'Conservationists are working ____ to remove snares and protect the saola\'s remaining habitat. (TIRE)', 'tirelessly', 'Cần trạng từ bổ nghĩa cho động từ "working". "Tire" -> "tireless" -> "tirelessly".'),
    textQ(33, 'Word Form', 'The ____ beauty of the saola has made it an important symbol for wildlife conservation in Vietnam. (NATURE)', 'natural', 'Cần tính từ bổ nghĩa cho "beauty". "Nature" -> "natural".'),
    textQ(34, 'Word Form', 'The ____ of forests for agriculture is one of the biggest threats to wildlife in Southeast Asia. (DESTROY)', 'destruction', 'Cần danh từ sau mạo từ "The". "Destroy" -> "destruction".'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "breed /briːd/ verb: to keep animals for the purpose of producing young animals." Complete: The zoo has successfully managed to ____ several endangered turtles in captivity this year.', 'breed', '"Breed" (nhân giống) phù hợp với định nghĩa và ngữ cảnh nhân giống rùa quý hiếm.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "rescue /ˈreskjuː/ verb: to save someone or something from a dangerous or harmful situation." Complete: The wildlife center works around the clock to ____ injured animals and nurse them back to health.', 'rescue', '"Rescue" (cứu hộ) phù hợp với định nghĩa và ngữ cảnh cứu động vật bị thương.'),
    textQ(37, 'Sentence Transformation', 'We must protect endangered species from illegal hunters and poachers. (PROTECTED)\n→ Endangered species ____ from illegal hunters and poachers.', 'must be protected', 'Câu bị động với động từ khuyết thiếu: S + must + be + V3/ed.'),
    textQ(38, 'Sentence Transformation', 'The population of wild tigers has decreased by over 95% in the last century. (DECREASE)\n→ There ____ in the population of wild tigers in the last century.', 'has been a decrease of over 95%', '"S + has decreased by X%" → "There has been a decrease of X% in + N".'),
    textQ(39, 'Sentence Transformation', 'Poachers kill rhinos for their horns, which are used in traditional medicine. (KILLED)\n→ Rhinos ____ for their horns, which are used in traditional medicine.', 'are killed by poachers', 'Câu bị động thì hiện tại đơn: S (số nhiều) + are + V3/ed + by O.'),
    textQ(40, 'Sentence Transformation', '"Do not buy products made from endangered animals," the guide told the tourists. (ADVISED)\n→ The guide ____ products made from endangered animals.', 'advised the tourists not to buy', '"Do not V," S told O → "S + advised + O + not to V".'),
  ];

  seedSingleTest_({
    test_id: '29',
    title: 'Đề thi thử số 29',
    description: 'Chủ đề Bảo vệ động vật quý hiếm - Từ vựng về động vật hoang dã',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

// ======================================================================
// ĐỀ 30: Tổng ôn tập cuối cùng - Tổng hợp tất cả chủ điểm trọng tâm
// ======================================================================
function seedQuestionsTest30_() {
  var clozePassage = [
    "Education is the foundation upon (17)____ a successful and prosperous society is built. In the 21st century,",
    "education systems around the world are evolving (18)____ meet the demands of a rapidly changing global economy.",
    "Skills such as critical thinking, creativity, collaboration, and digital literacy are now considered just as",
    "important (19)____ traditional academic knowledge. Vietnam has made significant progress in improving its",
    "education system, with increased investment in teacher training and school (20)____. However, challenges remain,",
    "(21)____ the need to reduce the gap between urban and rural education quality. Experts believe that the key to",
    "success lies in a balanced approach (22)____ combines academic excellence with practical skill development."
  ].join('\n\n');

  var readingPassage = [
    "In today's interconnected world, being bilingual or multilingual is an increasingly valuable asset. Research has",
    "shown that learning a second language not only improves communication abilities but also enhances cognitive",
    "functions such as memory, attention, and problem-solving. Bilingual individuals have been found to be better at",
    "multitasking and more resistant to age-related cognitive decline. In Vietnam, English is the most commonly learned",
    "foreign language, with millions of students studying it from primary school through university. The government's",
    "National Foreign Language Project aims to have all Vietnamese students achieve a basic level of English proficiency",
    "by the time they graduate from high school. However, achieving this goal has proven challenging, particularly in",
    "rural areas where there is a shortage of qualified English teachers. Language experts suggest that the most effective",
    "way to learn a language is through regular exposure and real communication, rather than just memorizing grammar rules.",
    "'Language is not just a school subject,' said one expert. 'It is a bridge that connects us to other cultures,'"
    + " 'and opens doors to opportunities around the world.'"
  ].join('\n\n');

  var qs = [
    mcq(1, 'Pronunciation: -ed endings', 'Choose the word whose underlined part is pronounced differently: A. invested B. improved C. reduced D. launched', 'A. invested /ɪd/', 'B. improved /d/', 'C. reduced /t/', 'D. launched /t/', 'A', '"Invested" có âm cuối /t/ nên -ed đọc /ɪd/, các từ còn lại đọc /d/ hoặc /t/.'),
    mcq(2, 'Pronunciation: Vowel sounds', 'Choose the word whose underlined part is pronounced differently: A. education B. foundation C. nation D. question', 'A. education /eɪ/', 'B. foundation /eɪ/', 'C. nation /eɪ/', 'D. question /tʃ/', 'D', '"Question" có âm "tion" đọc là /tʃən/, các từ còn lại có âm /ʃən/.'),
    mcq(3, 'Word Stress: 2-syllable', 'Choose the word with a different stress pattern: A. progress B. balance C. achieve D. system', 'A. progress', 'B. balance', 'C. achieve', 'D. system', 'C', '"Achieve" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(4, 'Word Stress: 3-syllable', 'Choose the word with a different stress pattern: A. government B. develop C. benefit D. heritage', 'A. government', 'B. develop', 'C. benefit', 'D. heritage', 'B', '"Develop" nhấn âm 2, các từ còn lại nhấn âm 1.'),
    mcq(5, 'Vocabulary in Context', 'The new curriculum aims to ____ students with both academic knowledge and practical skills for the modern workforce.', 'A. equip', 'B. deprive', 'C. remove', 'D. separate', 'A', '"Equip someone with something" = trang bị cho ai cái gì. Là cụm từ trong giáo dục.'),
    mcq(6, 'Grammar: Tenses', 'Vietnam ____ remarkable progress in expanding access to education over the past two decades.', 'A. made', 'B. has made', 'C. makes', 'D. is making', 'B', 'Thì hiện tại hoàn thành với "over the past two decades".'),
    mcq(7, 'Phrasal Verbs', 'The school decided to ____ a new after-school program to help struggling students improve their math skills.', 'A. set up', 'B. break down', 'C. put off', 'D. give away', 'A', '"Set up" = thiết lập, thành lập. "Set up a new program" = thiết lập chương trình mới.'),
    mcq(8, 'Prepositions', 'All students are expected to participate ____ at least one extracurricular activity each semester.', 'A. in', 'B. on', 'C. at', 'D. for', 'A', '"Participate in + N" = tham gia vào. Cấu trúc cố định.'),
    mcq(9, 'Grammar: Wish Clause', 'I wish I ____ more opportunities to practice speaking English with native speakers when I was younger.', 'A. had', 'B. have', 'C. will have', 'D. have had', 'A', 'Câu ước ở quá khứ: S + wish + S + had V3/ed. Hoặc hiện tại: had. Ngữ cảnh "when I was younger" → quá khứ → had had. Nhưng đáp án chỉ có had → ước hiện tại.'),
    mcq(10, 'Grammar: Conditional', 'If education systems ____ to the needs of the modern economy, students will be better prepared for future jobs.', 'A. adapt', 'B. adapted', 'C. will adapt', 'D. would adapt', 'A', 'Câu điều kiện loại 1: If + S + V (HTĐ), S + will + V.'),
    mcq(11, 'Vocabulary: Word Choice', 'Learning a second language has been shown to ____ cognitive functions such as memory and problem-solving skills.', 'A. enhance', 'B. reduce', 'C. prevent', 'D. ignore', 'A', '"Enhance" (tăng cường, nâng cao) là từ tích cực phù hợp với lợi ích của việc học ngôn ngữ.'),
    mcq(12, 'Communication', 'Student: "I am nervous about the final exam next week. Do you have any advice?"\nTeacher: "____"', 'A. Review your notes regularly, get plenty of rest, and believe in yourself. You have prepared well.', 'B. There is nothing you can do now. The exam is too close.', 'C. I do not care about your exam results.', 'D. Just copy from someone else during the test.', 'A', 'Lời khuyên tích cực, động viên từ giáo viên trước kỳ thi.'),
    mcq(13, 'Communication', 'Friend: "I am thinking about learning Japanese in addition to English. What do you think?"\nYou: "____"', 'A. That sounds like a great idea! Being trilingual will give you so many more opportunities in the future.', 'B. Learning languages is a waste of time. Just focus on one.', 'C. Japanese is too difficult. Nobody can learn it.', 'D. You should learn Chinese instead. Japanese is useless.', 'A', 'Phản hồi tích cực, khuyến khích bạn học thêm ngoại ngữ và nêu lợi ích.'),
    mcq(14, 'Public Signs', 'A sign in a library showing a finger over lips means:', 'A. Food and drinks are allowed', 'B. Please be quiet — silence is required', 'C. Group discussion area', 'D. Cell phone charging station', 'B', 'Biểu tượng ngón tay trên môi là dấu hiệu yêu cầu im lặng, phổ biến trong thư viện.'),
    mcq(15, 'Public Signs', 'What does a sign showing a graduation cap with a checkmark mean at a university?', 'A. Graduation ceremony registration', 'B. Degree verification / graduation requirements met', 'C. Cap and gown store', 'D. Student ID pickup', 'B', 'Biểu tượng mũ tốt nghiệp với dấu tick xác nhận đã hoàn thành yêu cầu tốt nghiệp.'),
    mcq(16, 'Grammar: Gerund/Infinitive', 'She refused ____ up even when the math problem seemed impossible to solve.', 'A. give', 'B. to give', 'C. giving', 'D. gave', 'B', '"Refuse + to V" là cấu trúc cố định.'),
    mcq(17, 'Cloze Test: Passage', clozePassage + '\n\nQ17: "Education is the foundation upon (17)____ a successful...society is built."', 'A. which', 'B. that', 'C. what', 'D. where', 'A', '"Upon which" trong mệnh đề quan hệ trang trọng thay cho "which...upon".'),
    mcq(18, 'Cloze Test', 'Fill in blank (18): "...are evolving (18)____ meet the demands of a rapidly changing global economy."', 'A. to', 'B. for', 'C. in', 'D. at', 'A', '"Evolve to V" = phát triển để làm gì. Chỉ mục đích.'),
    mcq(19, 'Cloze Test', 'Fill in blank (19): "...are now considered just as important (19)____ traditional academic knowledge."', 'A. as', 'B. than', 'C. like', 'D. with', 'A', 'So sánh ngang bằng: "as + adj + as".'),
    mcq(20, 'Cloze Test', 'Fill in blank (20): "...investment in teacher training and school (20)____."', 'A. infrastructure', 'B. instruction', 'C. instrument', 'D. institute', 'A', '"School infrastructure" = cơ sở hạ tầng trường học.'),
    mcq(21, 'Cloze Test', 'Fill in blank (21): "However, challenges remain, (21)____ the need to reduce the gap..."', 'A. including', 'B. includes', 'C. included', 'D. include', 'A', '"Including" (bao gồm) là giới từ bổ sung ý.'),
    mcq(22, 'Cloze Test', 'Fill in blank (22): "...a balanced approach (22)____ combines academic excellence with practical skill..."', 'A. that', 'B. who', 'C. what', 'D. where', 'A', '"That" thay cho "approach" trong mệnh đề quan hệ xác định.'),
    mcq(23, 'Reading Comprehension', readingPassage + '\n\nQ23: According to the passage, what is one cognitive benefit of learning a second language?', 'A. It makes people physically stronger', 'B. It enhances memory, attention, and problem-solving', 'C. It helps people eat more healthily', 'D. It improves eyesight', 'B', 'Đoạn văn nêu: "enhances cognitive functions such as memory, attention, and problem-solving".'),
    mcq(24, 'Reading: True/False', 'Q24: Bilingual people have been found to be worse at multitasking than monolingual people.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu người song ngữ "better at multitasking", không phải kém hơn.'),
    mcq(25, 'Reading: True/False', 'Q25: There is no shortage of qualified English teachers in rural Vietnam.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu có "a shortage of qualified English teachers" ở nông thôn.'),
    mcq(26, 'Reading: True/False', 'Q26: According to language experts, memorizing grammar rules is the most effective way to learn a language.', 'A. TRUE', '', 'B. FALSE', '', 'B', 'Đoạn văn nêu cách hiệu quả nhất là "regular exposure and real communication", không phải học vẹt ngữ pháp.'),
    mcq(27, 'Reading: MCQ', 'Q27: The word "cognitive" in the passage is closest in meaning to:', 'A. related to physical health', 'B. related to mental processes of thinking and learning', 'C. related to social relationships', 'D. related to financial matters', 'B', '"Cognitive" = liên quan đến nhận thức, quá trình tư duy và học tập của não bộ.'),
    mcq(28, 'Reading: Main Idea', 'Q28: What is the main message of this passage?', 'A. English is the only language worth learning', 'B. Learning a second language has many cognitive and practical benefits', 'C. Vietnamese students should stop learning foreign languages', 'D. Bilingual people are smarter than everyone else', 'B', 'Bài đọc nhấn mạnh nhiều lợi ích về nhận thức và thực tiễn của việc học ngôn ngữ thứ hai.'),
    textQ(29, 'Word Form', 'The ____ of technology in classrooms has fundamentally changed the way students learn and interact. (INTEGRATE)', 'integration', 'Cần danh từ sau mạo từ "The". "Integrate" (động từ) -> "integration" (danh từ).'),
    textQ(30, 'Word Form', 'Students who participate ____ in class discussions tend to perform better on exams. (ACTIVE)', 'actively', 'Cần trạng từ bổ nghĩa cho động từ "participate". "Active" -> "actively".'),
    textQ(31, 'Word Form', 'The ____ of lifelong learning is increasingly important in today\'s fast-changing job market. (IMPORTANT)', 'importance', 'Cần danh từ sau mạo từ "The". "Important" (tính từ) -> "importance" (danh từ).'),
    textQ(32, 'Word Form', 'Teachers play a ____ role in shaping the attitudes and values of the next generation. (CRUCIALITY)', 'crucial', 'Cần tính từ bổ nghĩa cho "role". "Cruciality" (danh từ) -> "crucial" (tính từ).'),
    textQ(33, 'Word Form', 'She managed her time ____, balancing schoolwork, extracurricular activities, and a part-time job. (EFFECT)', 'effectively', 'Cần trạng từ bổ nghĩa cho động từ "managed". "Effect" -> "effective" -> "effectively".'),
    textQ(34, 'Word Form', 'The National Foreign Language Project is an ____ initiative to improve English proficiency nationwide. (AMBITION)', 'ambitious', 'Cần tính từ bổ nghĩa cho "initiative". "Ambition" (danh từ) -> "ambitious" (tính từ).'),
    textQ(35, 'Dictionary Insertion', 'Look at the dictionary entry: "acquire /əˈkwaɪər/ verb: to gain or obtain something through effort or experience." Complete: Living abroad is one of the best ways to ____ a new language naturally.', 'acquire', '"Acquire a language" (tiếp thu ngôn ngữ) phù hợp với định nghĩa và ngữ cảnh học ngôn ngữ tự nhiên.'),
    textQ(36, 'Dictionary Insertion', 'Look at the dictionary entry: "achieve /əˈtʃiːv/ verb: to reach or attain a desired objective, level, or result by effort, skill, or courage." Complete: With hard work and determination, you can ____ any goal you set for yourself.', 'achieve', '"Achieve a goal" (đạt được mục tiêu) phù hợp với định nghĩa và ngữ cảnh động viên.'),
    textQ(37, 'Sentence Transformation', 'People believe that education is the key to reducing poverty worldwide. (BELIEVED)\n→ Education ____ the key to reducing poverty worldwide.', 'is believed to be', '"People believe that S + V" → "S + is/are believed + to V".'),
    textQ(38, 'Sentence Transformation', 'The last time I attended a formal English class was six months ago. (ATTENDED)\n→ I ____ a formal English class for six months.', 'have not attended', '"The last time + S + V (QKĐ) + was + time ago" → "S + have/has not V3/ed + for + time".'),
    textQ(39, 'Sentence Transformation', '"Study a little bit every day instead of cramming before the exam," the professor advised. (ADVISED)\n→ The professor ____ a little bit every day instead of cramming before the exam.', 'advised us to study', '"V," S advised O → "S + advised + O + to V".'),
    textQ(40, 'Sentence Transformation', 'If you do not review your lessons regularly, you will struggle during the final exam. (UNLESS)\n→ ____ your lessons regularly, you will struggle during the final exam.', 'Unless you review', '"If...not" → "Unless". Unless + S + V (HTĐ) = If + S + do/does not + V.'),
  ];

  seedSingleTest_({
    test_id: '30',
    title: 'Đề thi thử số 30',
    description: 'Tổng ôn tập cuối cùng - Tổng hợp tất cả chủ điểm trọng tâm',
    time_limit: 90,
    default_points: 0.25
  }, qs);
}

