// Seed two 2026 entrance-test datasets into the existing Tests/Questions sheets.
// Paste this file into Apps Script, keep Code.gs in the same project, then run:
// seedNew2026Tests()

function seedNew2026Tests() {
  seedKHTN2026_345_();
  seedDaNang2026_();
  Logger.log('Seeded new 2026 tests.');
}

function seedSingleTest_(testInfo, questions) {
  var testsSheet = getSheet_(SHEET_TESTS);
  var questionsSheet = getSheet_(SHEET_QUESTIONS);
  removeExistingTestRows_(testsSheet, questionsSheet, testInfo.test_id);

  testsSheet.appendRow([
    testInfo.test_id,
    testInfo.title,
    testInfo.description,
    testInfo.time_limit
  ]);

  if (!questions.length) return;

  var rows = questions.map(function(q, index) {
    return [
      testInfo.test_id + '_Q' + q.number,
      testInfo.test_id,
      q.part_number || 2,
      q.section_title || '',
      q.question_text || '',
      q.context_image_url || '',
      q.option_a || '',
      q.option_b || '',
      q.option_c || '',
      q.option_d || '',
      String(q.correct_answer || '').trim(),
      q.points || testInfo.default_points || CONFIG.POINTS_PER_QUESTION,
      q.explanation_template || ''
    ];
  });

  questionsSheet
    .getRange(questionsSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}

function removeExistingTestRows_(testsSheet, questionsSheet, testId) {
  deleteRowsByColumnValue_(questionsSheet, 2, testId);
  deleteRowsByColumnValue_(testsSheet, 1, testId);
}

function deleteRowsByColumnValue_(sheet, columnIndex, value) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  var values = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(value)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function seedKHTN2026_345_() {
  var passageFloraFauna = [
    'Plants and animals are an essential part of Earth\'s ecology. When we speak about a specific habitat and time, we refer to them as flora and fauna. Flora refers to plants and fauna refers to animals. They are fascinating to study due to their beauty and significance to human life.',
    'First, the flora and fauna on Earth create an ecological balance, making life possible for humans. Flora releases oxygen for humans. The carbon dioxide we breathe out is vital to plants. Humans also rely on the plants and animals for food, medicine, and more.',
    'Second, plants and animals combine to create a food chain. In the food chain, the animals play an important part. They feed on plants and, in turn, are the prey of other animals. Their droppings become fertilizers for plants.',
    'Finally, plants and animals add a lot of beauty to our Earth. People love and appreciate the natural world, but it can change when people damage habitats.'
  ].join('\\n\\n');

  var socialMedia = [
    'PROTECTING YOURSELF ON SOCIAL MEDIA',
    'Internet is Permanent: Once you (23) something online, it may never go away, even if you delete it.',
    'Be Selective with Friends: Only accept friend requests from people you know (24) to avoid fake profiles.',
    'Caution with Links: Be careful when clicking on links, even (25) friends. Be skeptical of offers that seem too good to be true.',
    'Manage Privacy Settings: Regularly check your privacy settings to ensure you\'re only sharing with (26) friends and family.',
    'Linked Information Across Networks: Be aware that sharing on one platform might also mean sharing on another, depending on your settings.'
  ].join('\\n');

  var overtourism = [
    'TOURISTS, GO HOME',
    'In Spain, there (27) a lot of news about overtourism recently. Overtourism happens when too many tourists visit a place, causing problems for local people and the environment. Many residents are unhappy because the high number of tourists makes their daily life difficult. The streets and public places become very crowded, and prices for things like food and rent (28). Thousands of people in Spain are protesting against mass tourism. In Barcelona, residents have sprayed tourists in restaurants with water guns. On the island of Mallorca, 10,000 people went to the streets holding banners that read "Enough is enough" and "Mallorca is not for sale."',
    'To solve this problem, the government is (29). They are trying to limit (30) tourists in some popular areas and are encouraging visitors to explore less-known places. (31) some new rules are being introduced to protect the environment and reduce the impact of tourism. Local businesses are also helping by promoting activities that respect the local culture and environment. For example, some tour companies are offering (32) tours and working to support local traditions. It is important to balance tourism with the needs of local communities. By making changes, Spain hopes to improve the situation for both residents and tourists.'
  ].join('\\n\\n');

  var chores = [
    'SHOULD KIDS DO CHORES?',
    'When you were a kid, did you help your parents around the house? Did you do the dishes, vacuum the floor, set the table, clean the bathroom, or do the laundry? (35) First of all, chores teach kids about responsibility. Life is not just fun and games! Sometimes, we do not want to do certain things, but we have to do them. You can give your child a simple task to teach this lesson. For example, you can ask them to make their bed every morning.',
    'House chores also teach kids to help their family. When they are older, your kids will have to help society. They will also have to work with other people. You can prepare them for this when you ask them to set the table before dinner, or when you ask them to clean up the living room before guests arrive. (36) but that they are part of a group.',
    'Finally, chores help kids to be active at home. In today\'s world, it is easy for kids to spend all their time in front of a screen playing video games, watching videos, or messaging their friends. Relaxation is important, but we must be careful (37)',
    'In conclusion, (38). They teach kids to be responsible. They teach them to help other people. And they teach them to be active around the home. A little housework can make a big difference!'
  ].join('\\n\\n');

  var qs = [
    mcq(1, 'Grammar and Vocabulary', 'The luggage was _____ than the doorman had exported.', 'mach heavy', 'heavier', 'more heavy', 'more heavier', 'B'),
    mcq(2, 'Grammar and Vocabulary', 'My mother asked me _____ when she called last week.', 'doing', 'what I were doing', 'was doing', 'am doing', 'C'),
    mcq(3, 'Grammar and Vocabulary', 'How _____ times have you watched that movie since its release?', 'much', 'few', 'little', 'many', 'D'),
    mcq(4, 'Grammar and Vocabulary', 'An artist does _____ tasks, such as painting or making music.', 'creative', 'create', 'creation', 'Creator', 'A'),
    mcq(5, 'Grammar and Vocabulary', 'My mom wondered _____ to pick up first, me or my younger brother.', 'that', 'who', 'when', 'which', 'B'),
    mcq(6, 'Grammar and Vocabulary', 'Among its many other _____ uses, fresh water is required by humans for drinking and agriculture.', 'agriculture', 'agriculturer', 'agriculturist', 'agricultural', 'A'),
    mcq(7, 'Grammar and Vocabulary', 'If he doesn\'t tell me the reason why he came late, I _____ be angry.', 'will', 'could', 'ought', 'will to', 'A'),
    mcq(8, 'Grammar and Vocabulary', 'Jack: "Let me show you how to use that library card."', 'I can do myself.', 'The librarian is not here.', 'No, I wouldn\'t.', 'That\'s very kind of you.', 'D'),
    mcq(9, 'Sentence Transformation', 'They decorated the tree with colorful Christmas balls.', 'They had decorated colorful Christmas balls and the tree.', 'Colorful Christmas balls had been decorated under the tree.', 'The tree had been decorated with colorful Christmas balls.', 'The tree was decorated with colorful Christmas balls.', 'D'),
    mcq(10, 'Sentence Transformation', '"Where did you go on vacation last summer?" I asked her.', 'I asked her where she went on vacation the previous summer.', 'I asked her where did she go on vacation the previous summer.', 'I asked her where she had gone on vacation the previous summer.', 'I asked her where had she gone on vacation the previous summer.', 'C'),
    mcq(11, 'Sentence Formation', 'We / hope / our country / switch to / sustainable fossil fuel alternatives / future.', 'We hope our country will switch to sustainable fossil fuel alternatives in the future.', 'We hope our country would switched to sustainable fossil fuel alternatives in the future.', 'We hope our country had switched to sustainable fossil fuel alternatives in the future.', 'We hope our country would have switched to sustainable fossil fuel alternatives in the future.', 'A'),
    mcq(12, 'Sentence Formation', 'He / advise / his friend / apologize / her mistake.', 'He advised his friend apologizing for her mistake.', 'He advised his friend that she will apologize for her mistake.', 'He advised his friend to have apologized for her mistake.', 'He advised his friend to apologize for her mistake.', 'D'),
    mcq(13, 'Reading Comprehension', passageFloraFauna + '\\n\\nQuestion 13 prompt is missing from the source document. Answer key says A.', '[Missing option A - answer key says A]', '[Missing option B]', '[Missing option C]', 'Humans do not depend on plants for oxygen.', 'A', 'Question/source data incomplete. Replace this row if you obtain the missing prompt/options.'),
    mcq(14, 'Reading Comprehension', passageFloraFauna + '\\n\\nWhich of the following is an example of the importance of flora?', 'Humans do not need plants.', 'Animals only eat other animals.', 'Plants produce oxygen for humans.', 'Plants destroy ecological balance.', 'C'),
    mcq(15, 'Reading Comprehension', passageFloraFauna + '\\n\\nThe word "vital" in paragraph 2 is CLOSEST in meaning to', 'useless', 'dangerous', 'necessary', 'interesting', 'C'),
    mcq(16, 'Reading Comprehension', passageFloraFauna + '\\n\\nAccording to the passage, what role do animals play in the food chain?', 'They eat plants and are eaten by other animals.', 'They help produce oxygen for plants.', 'They mainly eat other animals.', 'They have little effect on plants.', 'A'),
    mcq(17, 'Reading Comprehension', passageFloraFauna + '\\n\\nThe word "change" in the last paragraph is OPPOSITE in meaning to', 'increase', 'development', 'stability', 'difference', 'C'),
    mcq(18, 'Reading Comprehension', passageFloraFauna + '\\n\\nWhat is the main idea of the passage?', 'Plants and animals are beautiful but not important.', 'Only animals are important to humans.', 'Flora and fauna are essential to life and need protection.', 'Humans should stop using natural resources completely.', 'C'),
    mcq(19, 'Paragraph Cohesion', 'Put the sentences (a-c) in the correct order, then fill in the blank to make a logical text.\\nLast summer, I decided to find a part-time job to gain some work experience. I spent several days preparing my CV and looking for job advertisements online.\\na. Although I was nervous at first, I soon learned how to communicate with customers.\\nb. I finally got a job at a small café near my house.\\nc. By the end of the summer, I had saved some money and learned many useful skills.', 'b-a-c', 'a-b-c', 'c-a-b', 'a-c-b', 'A'),
    mcq(20, 'Paragraph Cohesion', 'Choose the sentence that can end the text in Question 19 most appropriately.', 'The experience helped me become more independent and responsible.', 'I forgot everything I learned from that job.', 'I decided never to work again because it was too tiring.', 'Working part-time is always a waste of time.', 'A'),
    mcq(21, 'Pronunciation', 'Choose the word whose underlined part differs from the other three in pronunciation.', 'healthy', 'meaning', 'bread', 'bear', 'B'),
    mcq(22, 'Pronunciation', 'Choose the word whose underlined part differs from the other three in pronunciation.', 'believes', 'deals', 'becomes', 'attracts', 'D'),
    mcq(23, 'Cloze Test 1: Social Media', socialMedia + '\\n\\nQuestion 23.', 'taking', 'share', 'look', 'share at', 'B'),
    mcq(24, 'Cloze Test 1: Social Media', socialMedia + '\\n\\nQuestion 24.', 'person', 'personality', 'personal', 'personally', 'D'),
    mcq(25, 'Cloze Test 1: Social Media', socialMedia + '\\n\\nQuestion 25.', 'from', 'to', 'over', 'in', 'A'),
    mcq(26, 'Cloze Test 1: Social Media', socialMedia + '\\n\\nQuestion 26.', 'a', 'the', 'no article', 'an', 'C'),
    mcq(27, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 27.', 'are', 'were', 'had been', 'has been', 'D'),
    mcq(28, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 28.', 'get up', 'go up', 'break up', 'use up', 'B'),
    mcq(29, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 29.', 'taking action', 'making action', 'doing action', 'drawing action', 'A'),
    mcq(30, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 30.', 'amount of', 'the number of', 'the amount of', 'number of', 'B'),
    mcq(31, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 31.', 'Conversely', 'Traditionally', 'In contrast', 'Additionally', 'D'),
    mcq(32, 'Cloze Test 2: Overtourism', overtourism + '\\n\\nQuestion 32.', 'eco-friendly', 'eco-friendliness', 'friend-eco', 'friendship-eco', 'A'),
    mcq(33, 'Stress', 'Choose the word that differs from the other three in the position of primary stress.', 'performance', 'tradition', 'heritage', 'relation', 'C'),
    mcq(34, 'Stress', 'Choose the word that differs from the other three in the position of primary stress.', 'answer', 'decide', 'suggest', 'require', 'A'),
    mcq(35, 'Paragraph Completion', chores + '\\n\\nQuestion 35.', 'chores teach kids important life skills and lessons', 'because too much relaxation can make people lazy', 'These are all important life skills, and they teach children useful life lessons', 'Kids need to learn that they are not kings and queens', 'C'),
    mcq(36, 'Paragraph Completion', chores + '\\n\\nQuestion 36.', 'chores teach kids important life skills and lessons', 'because too much relaxation can make people lazy', 'These are all important life skills, and they teach children useful life lessons', 'Kids need to learn that they are not kings and queens', 'D'),
    mcq(37, 'Paragraph Completion', chores + '\\n\\nQuestion 37.', 'chores teach kids important life skills and lessons', 'because too much relaxation can make people lazy', 'These are all important life skills, and they teach children useful life lessons', 'Kids need to learn that they are not kings and queens', 'B'),
    mcq(38, 'Paragraph Completion', chores + '\\n\\nQuestion 38.', 'chores teach kids important life skills and lessons', 'because too much relaxation can make people lazy', 'These are all important life skills, and they teach children useful life lessons', 'Kids need to learn that they are not kings and queens', 'A'),
    mcq(39, 'Reading Signs & Notices', 'What does the notice say?\\nNotice text: PARKING 2 hours maximum £4 per hour. Parking penalty charge notice: £50.', 'It costs at least £50 to park here for two hours.', 'You have to pay a fine if you park here for three hours.', 'There are no parking charges after 2.00 p.m.', 'If you park for an hour, you have to pay £54.', 'B'),
    mcq(40, 'Reading Signs & Notices', 'What does the sign say?\\nSign text: SILENCE! exams in progress.', 'Do not walk into the exam area.', 'Do not look at other students.', 'Do not tell questions to the examiner.', 'Please remain quiet during the examination.', 'D')
  ];

  seedSingleTest_({
    test_id: 'KHTN2026_345',
    title: 'Đề thi tuyển sinh vào lớp 10 THPT Chuyên KHTN 2026 - Mã đề 345',
    description: 'ĐHQG Hà Nội - Trường Đại học Khoa học Tự nhiên. Môn Tiếng Anh.',
    time_limit: 60,
    default_points: 0.25
  }, qs);
}

function seedDaNang2026_() {
  var passageWright = [
    'The inventors of the first airplane were not scientists. In fact, they did not even finish high school. The inventors of the first airplane were just ordinary brothers with a big dream - a dream of flying.',
    'Wilbur and Orville Wright grew up in Indiana in a large family. Neither of the brothers finished high school. Instead of sitting in class and reading, they wanted to work and make things, like machines. When the brothers were 18 and 22 years old, they started up their own printing press and bicycle shop.',
    'One day, Wilbur read an interesting story in the newspaper about a man trying to fly in a glider. The man died, but the story gave Wilbur an idea. He decided to teach himself about flying in order to make the perfect glider. Together, Orville and Wilbur successfully tested their new glider on the beach. Then the brothers decided to make their glider into a flying machine. They put an engine and propeller on the glider and called their new machine the "Wright Flyer". The plane broke down on its first test, but the Wright brothers didn\'t give up.',
    'With a little more hard work, the Wright Flyer made a successful flight. On December 17, 1903, the Wright brothers made history by flying the Wright Flyer over the beach at Kitty Hawk, North Carolina. The airplane was born.'
  ].join('\\n\\n');

  var qs = [
    mcq(1, 'I. Multiple Choice', 'I wish I _____ fluent in Korean to communicate with my Kpop idols.', 'am', 'will be', 'were', 'be', 'C'),
    mcq(2, 'I. Multiple Choice', 'How _____ sugar is there in a glass of milk tea?', 'any', 'few', 'many', 'much', 'D'),
    mcq(3, 'I. Multiple Choice', 'The town\'s suitable public transportation makes it a _____ place for its residents.', 'dangerous', 'liveable', 'remote', 'confident', 'B'),
    mcq(4, 'I. Multiple Choice', 'The children couldn\'t come to class on time yesterday afternoon _____ it rained heavily.', 'in spite of', 'although', 'because', 'because of', 'C'),
    mcq(5, 'I. Multiple Choice', 'What should we do to reduce levels of environmental _____?', 'pollution', 'construction', 'population', 'congestion', 'A'),
    mcq(6, 'I. Multiple Choice', 'If the villagers promote their handicrafts well, more tourists _____ their products.', 'buys', 'will buy', 'bought', 'were buying', 'B'),
    mcq(7, 'I. Multiple Choice', 'Please let me know _____ to put these new words.', 'why', 'where', 'whose', 'which', 'B'),
    mcq(8, 'I. Multiple Choice', 'As the eldest daughter in the family, Emily has to _____ her younger brothers.', 'look after', 'pass down', 'go over', 'apply for', 'A'),
    mcq(9, 'I. Multiple Choice', 'Daisy _____ a good impression on my parents. She was so polite.', 'played', 'made', 'built', 'caused', 'B'),
    mcq(10, 'I. Multiple Choice', 'Due to climate change, this summer is _____ than the previous one.', 'the hottest', 'as hot', 'so hot', 'hotter', 'D'),
    mcq(11, 'I. Multiple Choice', 'John: "I\'ll have the final exam tomorrow." Tom: "_____"', 'Good luck!', 'Not at all.', 'Great idea.', 'My pleasure.', 'A'),
    mcq(12, 'I. Multiple Choice', 'Arrange the dialogue. a. Sure! It\'s about a 20-minute walk from here. Just head straight down this road and turn left at the second traffic light. b. Excuse me, could you show me the way to the nearest train station? I\'m lost. c. Thank you! And will I see any signs along the way? d. Got it! Thanks a lot for your help! e. Yes, there are clear signs pointing to the station once you make that turn.', 'b - a - e - d - c', 'b - a - c - e - d', 'b - a - c - d - e', 'b - a - d - c - e', 'C'),
    textQ(13, 'II. Word Form', 'We need a good antivirus software to _____ our computers. (protection)', 'protect'),
    textQ(14, 'II. Word Form', 'It is necessary for teenagers to eat fruits and vegetables _____. (regular)', 'regularly'),
    textQ(15, 'II. Word Form', 'Tet is an important _____ of Vietnamese people. (celebrate)', 'celebration'),
    textQ(16, 'II. Word Form', 'These students come from _____ countries in Asia. (difference)', 'different'),
    textQ(17, 'III. Verb Form', 'Helen usually _____ a shower after she has breakfast. (take)', 'takes'),
    textQ(18, 'III. Verb Form', 'They decided _____ up a new company two months ago. (set)', 'to set'),
    textQ(19, 'III. Verb Form', 'At 6 o\'clock yesterday, my grandfather _____ his old bike. (repair)', 'was repairing'),
    textQ(20, 'III. Verb Form', 'I suggest _____ a three-day trip to some historical places in Viet Nam. (have)', 'having'),
    mcq(21, 'IV. Reading Gap', 'Her journey will start in Tokyo _____ Italy.', 'on', 'in', 'at', 'for', 'B'),
    mcq(22, 'IV. Reading Gap', 'Sammy is a very _____ girl and is not worried about travelling alone.', 'dependent', 'independende', 'independent', 'dependently', 'C'),
    mcq(23, 'IV. Reading Gap', 'She doesn\'t want to cycle in darkness, _____ at the end of each afternoon.', 'so', 'nor', 'but', 'yet', 'A'),
    mcq(24, 'IV. Reading Gap', 'The journey will take about 40 days and she hopes to receive lots of _____ on her fund-raising web page.', 'prices', 'donations', 'rumours', 'warnings', 'B'),
    mcq(25, 'V. Signs and Notices', 'Important notice: Class 9A Music moves to Room 6 today. Starting at 2.30 a.m., as usual.', 'There is no room for the music lesson today.', 'There is no music lesson today.', 'The music lesson begins later today.', 'The music class is in another room today.', 'D'),
    mcq(26, 'V. Signs and Notices', 'Children\'s play area. No dogs allowed.', 'Only dogs can play in this area.', 'Dogs can\'t enter the children\'s play area.', 'Dogs are safe for children in this play area.', 'Don\'t let your children play near dogs.', 'B'),
    mcq(27, 'V. Signs and Notices', 'Library hours: open 9 a.m. till 5.30 p.m. weekdays; 9.30 a.m. till 1 p.m. Saturday.', 'The library is not open on Sunday.', 'The library is open every day.', 'The library on Saturday only.', 'The library not open on weekdays.', 'A'),
    mcq(28, 'V. Signs and Notices', 'No smoking.', 'Only adults can smoke here.', 'The smoking zone isn\'t working.', 'You mustn\'t smoke in this area.', 'Smoking is acceptable right here.', 'C'),
    mcq(29, 'VI. Reading Passage', passageWright + '\\n\\nQuestion 1.', 'The Wright brothers were scientists.', 'How the Wright brothers made a first plane.', 'The history of high school education.', 'Why gliders cannot fly.', 'B'),
    mcq(30, 'VI. Reading Passage', passageWright + '\\n\\nQuestion 2.', 'They owned a bike shop.', 'They were university scientists.', 'They worked in a library.', 'They were pilots in North Carolina.', 'A'),
    mcq(31, 'VI. Reading Passage', passageWright + '\\n\\nQuestion 3. The word "ordinary" is closest in meaning to', 'famous', 'special', 'normal', 'modern', 'C'),
    mcq(32, 'VI. Reading Passage', passageWright + '\\n\\nQuestion 4. Source image only provided answer D; replace options if you have the full source.', 'Missing option A', 'Missing option B', 'Missing option C', 'Missing option D', 'D', 'Question/source data incomplete.'),
    textQ(33, 'VII. Sentence Transformation', '"Do you like playing the guitar, son?" asked Mrs. Kim. -> Mrs. Kim asked her son _____ playing the guitar.', 'if he liked'),
    textQ(34, 'VII. Sentence Transformation', 'Despite her contribution to the project, she received no recognition from her boss. -> Although she _____, she received no recognition from her boss.', 'had contributed to the project'),
    textQ(35, 'VIII. Sentence Combination', 'Steven was sleeping. Lucy didn\'t want to wake him up. (while) -> Lucy ____________________ sleeping.', 'did not want to wake Steven up while he was'),
    textQ(36, 'VIII. Sentence Combination', 'The man is an experienced English teacher. He helped me with my pronunciation. (who) -> The man ____________________ English teacher.', 'who helped me with my pronunciation is an experienced')
  ];

  seedSingleTest_({
    test_id: 'DANANG2026',
    title: 'Đề thi tuyển sinh vào lớp 10 Đà Nẵng 2026-2027',
    description: 'Sở GD&ĐT Đà Nẵng. Môn Tiếng Anh. Source includes 36 available items, including text-answer items.',
    time_limit: 60,
    default_points: 10 / 36
  }, qs);
}

function mcq(number, section, text, a, b, c, d, answer, explanation) {
  return {
    number: number,
    part_number: 2,
    section_title: section,
    question_text: text,
    option_a: a,
    option_b: b,
    option_c: c,
    option_d: d,
    correct_answer: answer,
    explanation_template: explanation || ''
  };
}

function textQ(number, section, text, answer, explanation) {
  return {
    number: number,
    part_number: 4,
    section_title: section,
    question_text: text,
    correct_answer: answer,
    explanation_template: explanation || ''
  };
}
