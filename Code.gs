// ==========================================
// HCMC Grade 10 English Mock Test System
// Google Apps Script Backend - Code.gs
// ==========================================

// ---------- CONFIGURATION ----------
var CONFIG = {
  SPREADSHEET_ID: '1J6cnzgyFCvXfeHS9pNlzNQUaQb61g7fo_I1DQlUhSrI',
  DEEPSEEK_API_URL: 'https://api.deepseek.com/v1/chat/completions',
  DEEPSEEK_MODEL: 'deepseek-chat',
  TIME_LIMIT_MINUTES: 90,
  POINTS_PER_QUESTION: 0.25,
  MAX_SCORE: 10.0,
  TOTAL_QUESTIONS: 40
};

// Sheet names
var SHEET_TESTS = 'Tests';
var SHEET_QUESTIONS = 'Questions';
var SHEET_SUBMISSIONS = 'Submissions';

// Column indices (1-based for GAS)
var COL_TEST = { test_id: 1, title: 2, description: 3, time_limit: 4 };

var COL_QUESTION = {
  question_id: 1, test_id: 2, part_number: 3, section_title: 4,
  question_text: 5, context_image_url: 6, option_a: 7, option_b: 8,
  option_c: 9, option_d: 10, correct_answer: 11, points: 12,
  explanation_template: 13
};

var COL_SUBMISSION = {
  submission_id: 1, student_name: 2, student_id: 3,
  test_id: 4, answers_json: 5, score: 6, timestamp: 7
};

// ---------- WEB APP ENTRY POINT ----------
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e);
  }

  if (e && e.parameter && e.parameter.view === 'parent') {
    return getParentTemplate_()
      .evaluate()
      .setTitle('Parent View - Luyen thi vao 10 TPHCM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  return getIndexTemplate_()
    .evaluate()
    .setTitle('Luyen thi vao 10 TPHCM - Tieng Anh')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getIndexTemplate_() {
  try {
    return HtmlService.createTemplateFromFile('index');
  } catch (err) {
    return HtmlService.createTemplateFromFile('Index');
  }
}

function getParentTemplate_() {
  try {
    return HtmlService.createTemplateFromFile('parent');
  } catch (err) {
    return HtmlService.createTemplateFromFile('Parent');
  }
}

// ---------- STATIC SITE API ENTRY POINT ----------
function handleApiRequest_(e) {
  var callback = e.parameter.callback || '';
  var action = e.parameter.action || '';
  var payload = parsePayload_(e.parameter.payload);
  var result;

  try {
    switch (action) {
      case 'getTests':
        result = getTests();
        break;
      case 'getTestDetails':
        result = getTestDetails(payload);
        break;
      case 'getParentDashboard':
        result = getParentDashboard(payload || {});
        break;
      case 'getParentSubmissionDetail':
        result = getParentSubmissionDetail(payload || {});
        break;
      case 'submitTest':
        result = submitTest(payload || {});
        break;
      case 'callDeepSeekExplanation':
        result = callDeepSeekExplanation(payload || {});
        break;
      default:
        result = { error: true, message: 'Unknown API action: ' + action };
    }
  } catch (err) {
    result = { error: true, message: 'Server error: ' + err.toString() };
  }

  return createApiResponse_(result, callback);
}

function parsePayload_(payload) {
  if (!payload) return null;
  try {
    return JSON.parse(payload);
  } catch (err) {
    return payload;
  }
}

function createApiResponse_(data, callback) {
  var json = JSON.stringify(data);
  var output = callback
    ? callback.replace(/[^\w.$]/g, '') + '(' + json + ');'
    : json;
  var mimeType = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;

  return ContentService
    .createTextOutput(output)
    .setMimeType(mimeType);
}

function formatTimestamp_(value) {
  if (!value) return '';
  var date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  } catch (err) {
    return date.toISOString();
  }
}

function buildTestTitleMap_() {
  var map = {};
  var testsSheet = getSheet_(SHEET_TESTS);
  var testsData = testsSheet.getDataRange().getValues();
  for (var i = 1; i < testsData.length; i++) {
    var row = testsData[i];
    if (!row[COL_TEST.test_id - 1]) continue;
    map[String(row[COL_TEST.test_id - 1])] = row[COL_TEST.title - 1] || '';
  }
  return map;
}

function getParentDashboard(payload) {
  try {
    var limit = parseInt((payload && payload.limit) || 10, 10);
    if (isNaN(limit) || limit <= 0) limit = 10;
    if (limit > 50) limit = 50;

    var testTitleMap = buildTestTitleMap_();
    var submissionsSheet = getSheet_(SHEET_SUBMISSIONS);
    var data = submissionsSheet.getDataRange().getValues();
    var submissions = [];
    var byTest = {};
    var uniqueStudents = {};
    var scoreSum = 0;
    var latest = null;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[COL_SUBMISSION.submission_id - 1]) continue;

      var submissionId = row[COL_SUBMISSION.submission_id - 1];
      var studentName = row[COL_SUBMISSION.student_name - 1] || 'An danh';
      var studentId = row[COL_SUBMISSION.student_id - 1] || '';
      var testId = row[COL_SUBMISSION.test_id - 1];
      var score = parseFloat(row[COL_SUBMISSION.score - 1]) || 0;
      var timestampRaw = row[COL_SUBMISSION.timestamp - 1] || '';
      var timestampDate = new Date(timestampRaw);
      var timestampValue = isNaN(timestampDate.getTime()) ? 0 : timestampDate.getTime();
      var timestampDisplay = formatTimestamp_(timestampRaw);
      var testKey = String(testId);
      var testTitle = testTitleMap[testKey] || ('De thi ' + testKey);
      var studentKey = studentId ? (studentId + '|' + studentName) : studentName;

      var summaryItem = {
        submission_id: submissionId,
        student_name: studentName,
        student_id: studentId,
        test_id: testId,
        test_title: testTitle,
        score: score,
        timestamp: timestampRaw,
        timestamp_display: timestampDisplay,
        timestamp_value: timestampValue
      };
      submissions.push(summaryItem);
      scoreSum += score;
      uniqueStudents[studentKey] = true;

      if (!byTest[testKey]) {
        byTest[testKey] = {
          test_id: testId,
          test_title: testTitle,
          submission_count: 0,
          score_sum: 0,
          latest_timestamp: 0,
          latest_timestamp_display: '',
          latest_score: 0
        };
      }
      byTest[testKey].submission_count += 1;
      byTest[testKey].score_sum += score;
      if (timestampValue >= byTest[testKey].latest_timestamp) {
        byTest[testKey].latest_timestamp = timestampValue;
        byTest[testKey].latest_timestamp_display = timestampDisplay;
        byTest[testKey].latest_score = score;
      }

      if (!latest || timestampValue > latest.timestampValue) {
        latest = { timestampValue: timestampValue, timestampDisplay: timestampDisplay };
      }
    }

    submissions.sort(function(a, b) {
      return (b.timestamp_value || 0) - (a.timestamp_value || 0);
    });

    var testsSummary = Object.keys(byTest).map(function(key) {
      var item = byTest[key];
      var avg = item.submission_count ? item.score_sum / item.submission_count : 0;
      return {
        test_id: item.test_id,
        test_title: item.test_title,
        submission_count: item.submission_count,
        average_score: Math.round(avg * 100) / 100,
        latest_score: item.latest_score,
        latest_timestamp_display: item.latest_timestamp_display
      };
    }).sort(function(a, b) {
      return b.submission_count - a.submission_count;
    });

    var totalSubmissions = submissions.length;
    var averageScore = totalSubmissions ? scoreSum / totalSubmissions : 0;

    return {
      stats: {
        total_submissions: totalSubmissions,
        unique_students: Object.keys(uniqueStudents).length,
        average_score: Math.round(averageScore * 100) / 100,
        latest_activity: latest ? latest.timestampDisplay : ''
      },
      recent_submissions: submissions.slice(0, limit),
      test_summary: testsSummary,
      limit: limit
    };
  } catch (e) {
    return { error: true, message: 'Khong the tai du lieu phu huynh. Vui long thu lai.' };
  }
}

function getParentSubmissionDetail(payload) {
  try {
    var submissionId = '';
    if (typeof payload === 'string') {
      submissionId = payload;
    } else if (payload) {
      submissionId = payload.submission_id || payload.submissionId || '';
    }
    submissionId = String(submissionId || '').trim();
    if (!submissionId) {
      return { error: true, message: 'Thieu submission_id.' };
    }

    var submissionsSheet = getSheet_(SHEET_SUBMISSIONS);
    var submissionsData = submissionsSheet.getDataRange().getValues();
    var submissionRow = null;

    for (var i = 1; i < submissionsData.length; i++) {
      if (String(submissionsData[i][COL_SUBMISSION.submission_id - 1]) === submissionId) {
        submissionRow = submissionsData[i];
        break;
      }
    }

    if (!submissionRow) {
      return { error: true, message: 'Khong tim thay bai nop nay.' };
    }

    var testId = submissionRow[COL_SUBMISSION.test_id - 1];
    var testTitleMap = buildTestTitleMap_();
    var testTitle = testTitleMap[String(testId)] || ('De thi ' + testId);
    var studentName = submissionRow[COL_SUBMISSION.student_name - 1] || 'An danh';
    var studentId = submissionRow[COL_SUBMISSION.student_id - 1] || '';
    var score = parseFloat(submissionRow[COL_SUBMISSION.score - 1]) || 0;
    var timestamp = submissionRow[COL_SUBMISSION.timestamp - 1] || '';
    var answers = {};
    try {
      answers = JSON.parse(submissionRow[COL_SUBMISSION.answers_json - 1] || '{}') || {};
    } catch (err) {
      answers = {};
    }

    var questionsSheet = getSheet_(SHEET_QUESTIONS);
    var questionsData = questionsSheet.getDataRange().getValues();
    var questions = [];
    var correctCount = 0;
    var qDisplayNum = 0;

    for (var j = 1; j < questionsData.length; j++) {
      var qRow = questionsData[j];
      if (String(qRow[COL_QUESTION.test_id - 1]) !== String(testId)) continue;

      qDisplayNum++;
      var qId = qRow[COL_QUESTION.question_id - 1];
      var studentAnswer = String(answers[qId] || '').trim().toUpperCase();
      var correctAnswer = String(qRow[COL_QUESTION.correct_answer - 1] || '').trim().toUpperCase();
      var isCorrect = studentAnswer === correctAnswer;
      if (isCorrect) correctCount++;

      questions.push({
        question_id: qId,
        display_number: qDisplayNum,
        part_number: qRow[COL_QUESTION.part_number - 1],
        section_title: qRow[COL_QUESTION.section_title - 1] || '',
        question_text: qRow[COL_QUESTION.question_text - 1] || '',
        student_answer: studentAnswer || '(Chua tra loi)',
        correct_answer: correctAnswer,
        is_correct: isCorrect,
        points_earned: isCorrect ? (parseFloat(qRow[COL_QUESTION.points - 1]) || CONFIG.POINTS_PER_QUESTION) : 0,
        points_max: parseFloat(qRow[COL_QUESTION.points - 1]) || CONFIG.POINTS_PER_QUESTION
      });
    }

    return {
      submission: {
        submission_id: submissionId,
        student_name: studentName,
        student_id: studentId,
        test_id: testId,
        test_title: testTitle,
        score: score,
        max_score: CONFIG.MAX_SCORE,
        timestamp: timestamp,
        timestamp_display: formatTimestamp_(timestamp),
        total_questions: questions.length,
        correct_count: correctCount
      },
      questions: questions
    };
  } catch (e) {
    return { error: true, message: 'Khong the tai chi tiet bai nop. Vui long thu lai.' };
  }
}

// ---------- HELPER: Get or create sheet with headers ----------
function getSheet_(name) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    switch (name) {
      case SHEET_TESTS:
        sheet.appendRow(['test_id', 'title', 'description', 'time_limit']);
        break;
      case SHEET_QUESTIONS:
        sheet.appendRow(['question_id', 'test_id', 'part_number', 'section_title',
          'question_text', 'context_image_url', 'option_a', 'option_b',
          'option_c', 'option_d', 'correct_answer', 'points', 'explanation_template']);
        break;
      case SHEET_SUBMISSIONS:
        sheet.appendRow(['submission_id', 'student_name', 'student_id',
          'test_id', 'answers_json', 'score', 'timestamp']);
        break;
    }
  }
  return sheet;
}

// ---------- HELPER: Get DeepSeek API key from script properties ----------
function getApiKey_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('DEEPSEEK_API_KEY') || '';
}

// ---------- API: Get all tests ----------
function getTests() {
  try {
    var sheet = getSheet_(SHEET_TESTS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var tests = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[COL_TEST.test_id - 1]) continue;
      tests.push({
        test_id: row[COL_TEST.test_id - 1],
        title: row[COL_TEST.title - 1] || '',
        description: row[COL_TEST.description - 1] || '',
        time_limit: row[COL_TEST.time_limit - 1] || CONFIG.TIME_LIMIT_MINUTES
      });
    }
    return tests;
  } catch (e) {
    return { error: true, message: 'Khong the tai danh sach de thi. Vui long thu lai.' };
  }
}

// ---------- API: Get test details (WITHOUT correct answers) ----------
function getTestDetails(testId) {
  try {
    var sheet = getSheet_(SHEET_QUESTIONS);
    var data = sheet.getDataRange().getValues();

    // Get test info
    var testsSheet = getSheet_(SHEET_TESTS);
    var testsData = testsSheet.getDataRange().getValues();
    var testInfo = null;
    for (var i = 1; i < testsData.length; i++) {
      if (testsData[i][COL_TEST.test_id - 1] == testId) {
        testInfo = {
          test_id: testsData[i][COL_TEST.test_id - 1],
          title: testsData[i][COL_TEST.title - 1] || '',
          description: testsData[i][COL_TEST.description - 1] || '',
          time_limit: testsData[i][COL_TEST.time_limit - 1] || CONFIG.TIME_LIMIT_MINUTES
        };
        break;
      }
    }

    if (!testInfo) return { error: true, message: 'Khong tim thay de thi nay.' };

    // Extract questions for this test (exclude correct_answer)
    var questions = [];
    var qNum = 0;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[COL_QUESTION.test_id - 1] == testId) {
        qNum++;
        questions.push({
          question_id: row[COL_QUESTION.question_id - 1],
          test_id: row[COL_QUESTION.test_id - 1],
          display_number: qNum,
          part_number: row[COL_QUESTION.part_number - 1],
          section_title: row[COL_QUESTION.section_title - 1] || '',
          question_text: row[COL_QUESTION.question_text - 1] || '',
          context_image_url: row[COL_QUESTION.context_image_url - 1] || '',
          option_a: row[COL_QUESTION.option_a - 1] || '',
          option_b: row[COL_QUESTION.option_b - 1] || '',
          option_c: row[COL_QUESTION.option_c - 1] || '',
          option_d: row[COL_QUESTION.option_d - 1] || '',
          points: row[COL_QUESTION.points - 1] || CONFIG.POINTS_PER_QUESTION
        });
      }
    }

    return { test_info: testInfo, questions: questions };
  } catch (e) {
    return { error: true, message: 'Khong the tai noi dung de thi. Vui long thu lai.' };
  }
}

// ---------- API: Submit test and calculate score ----------
function submitTest(submissionData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var sheet = getSheet_(SHEET_QUESTIONS);
    var data = sheet.getDataRange().getValues();

    var testId = submissionData.test_id;
    var studentAnswers = submissionData.answers || {};
    var studentName = submissionData.student_name || 'An danh';
    var studentId = submissionData.student_id || '';

    // Build answer key and question detail maps from sheet
    var answerKey = {};
    var questionDetails = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[COL_QUESTION.test_id - 1] == testId) {
        var qId = row[COL_QUESTION.question_id - 1];
        answerKey[qId] = String(row[COL_QUESTION.correct_answer - 1] || '').trim().toUpperCase();
        questionDetails[qId] = {
          question_id: qId,
          part_number: row[COL_QUESTION.part_number - 1],
          section_title: row[COL_QUESTION.section_title - 1] || '',
          question_text: row[COL_QUESTION.question_text - 1] || '',
          option_a: row[COL_QUESTION.option_a - 1] || '',
          option_b: row[COL_QUESTION.option_b - 1] || '',
          option_c: row[COL_QUESTION.option_c - 1] || '',
          option_d: row[COL_QUESTION.option_d - 1] || '',
          correct_answer: String(row[COL_QUESTION.correct_answer - 1] || '').trim().toUpperCase(),
          points: row[COL_QUESTION.points - 1] || CONFIG.POINTS_PER_QUESTION,
          explanation_template: row[COL_QUESTION.explanation_template - 1] || ''
        };
      }
    }

    // Calculate scores
    var totalScore = 0;
    var results = [];
    var sectionScores = {
      '1': { correct: 0, total: 0, title: 'Phonetics', maxPoints: 0, earnedPoints: 0 },
      '2': { correct: 0, total: 0, title: 'Vocabulary, Grammar & Communication', maxPoints: 0, earnedPoints: 0 },
      '3': { correct: 0, total: 0, title: 'Reading Comprehension', maxPoints: 0, earnedPoints: 0 },
      '4': { correct: 0, total: 0, title: 'Writing', maxPoints: 0, earnedPoints: 0 }
    };

    var qDisplayNum = 0;
    for (var qId in questionDetails) {
      qDisplayNum++;
      var q = questionDetails[qId];
      var studentAnswer = String(studentAnswers[qId] || '').trim().toUpperCase();
      var correctAnswer = q.correct_answer;
      var isCorrect = studentAnswer === correctAnswer;

      if (isCorrect) {
        totalScore += q.points;
        sectionScores[String(q.part_number)].correct++;
        sectionScores[String(q.part_number)].earnedPoints += q.points;
      }
      sectionScores[String(q.part_number)].total++;
      sectionScores[String(q.part_number)].maxPoints += q.points;

      results.push({
        question_id: qId,
        display_number: qDisplayNum,
        part_number: q.part_number,
        section_title: q.section_title,
        question_text: q.question_text,
        option_a: q.option_a,
        option_b: q.option_b,
        option_c: q.option_c,
        option_d: q.option_d,
        student_answer: studentAnswer || '(Chua tra loi)',
        correct_answer: correctAnswer,
        is_correct: isCorrect,
        points_earned: isCorrect ? q.points : 0,
        points_max: q.points,
        explanation_template: q.explanation_template
      });
    }

    totalScore = Math.round(totalScore * 100) / 100;

    // Store submission
    var submissionsSheet = getSheet_(SHEET_SUBMISSIONS);
    var submissionId = 'SUB-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    submissionsSheet.appendRow([
      submissionId, studentName, studentId, testId,
      JSON.stringify(studentAnswers), totalScore, new Date().toISOString()
    ]);

    lock.releaseLock();

    return {
      submission_id: submissionId,
      test_id: testId,
      student_name: studentName,
      total_score: totalScore,
      max_score: CONFIG.MAX_SCORE,
      total_questions: results.length,
      correct_count: results.filter(function(r) { return r.is_correct; }).length,
      section_scores: sectionScores,
      results: results
    };
  } catch (e) {
    try { lock.releaseLock(); } catch (_) {}
    return { error: true, message: 'Khong the nop bai. Vui long thu lai. Chi tiet: ' + e.toString() };
  }
}

// ---------- API: Call DeepSeek for AI explanation ----------
function callDeepSeekExplanation(questionData) {
  try {
    var apiKey = getApiKey_();
    if (!apiKey) {
      return { error: true, message: 'DeepSeek API key chua duoc cau hinh. Vui long lien he quan tri vien.' };
    }

    var questionText = questionData.question_text || '';
    var studentAnswer = questionData.student_answer || '(Khong tra loi)';
    var correctAnswer = questionData.correct_answer || '';
    var customQuestion = String(questionData.custom_question || '').trim();
    var optionLabels = [];
    if (questionData.option_a) optionLabels.push('A. ' + questionData.option_a);
    if (questionData.option_b) optionLabels.push('B. ' + questionData.option_b);
    if (questionData.option_c) optionLabels.push('C. ' + questionData.option_c);
    if (questionData.option_d) optionLabels.push('D. ' + questionData.option_d);
    var optionsText = optionLabels.join('\n');

    var systemPrompt = 'Em la mot giao vien tieng Anh than thien, giau kinh nghiem luyen thi vao lop 10 tai TP.HCM. Nhiem vu cua em la giai thich can ke, de hieu bang tieng Viet cho hoc sinh.\n\n' +
      'Khi giai thich mot cau hoi trac nghiem tieng Anh, em hay:\n' +
      '0. Neu hoc sinh co cau hoi rieng ve cau nay, hay tra loi cau hoi do truoc mot cach ro rang, sau do moi giai thich cau goc.\n' +
      '1. Dich nghia cau hoi va cac dap an sang tieng Viet.\n' +
      '2. Phan tich ngu phap hoac tu vung lien quan den cau hoi.\n' +
      '3. Giai thich tai sao dap an dung (' + correctAnswer + ') la chinh xac.\n' +
      '4. Giai thich tai sao cac dap an con lai la sai.\n' +
      '5. Dua ra meo lam bai hoac kien thuc can ghi nho (neu co).\n\n' +
      'Hay tra loi voi giong dieu dong vien, khich le hoc sinh. Su dung tieng Viet hoan toan, chi giu nguyen cac thuat ngu tieng Anh chuyen nganh. ' +
      'Dinh dang cau tra loi bang HTML don gian (dung <p>, <strong>, <em>, <ul>, <li>) de hien thi dep tren web.';

    var userMessage = 'Cau hoi: ' + questionText + '\n\n' +
      (customQuestion ? 'Cau hoi rieng cua hoc sinh: ' + customQuestion + '\n\n' : '') +
      (optionsText ? 'Cac dap an:\n' + optionsText + '\n\n' : '') +
      'Dap an hoc sinh chon: ' + studentAnswer + '\n' +
      'Dap an dung: ' + correctAnswer + '\n\n' +
      'Em hay giai thich chi tiet cau hoi nay giup hoc sinh nhe!';

    var payload = {
      model: CONFIG.DEEPSEEK_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMessage }
      ],
      temperature: 0.7
    };

    var response = UrlFetchApp.fetch(CONFIG.DEEPSEEK_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode === 200) {
      var json = JSON.parse(responseText);
      var explanation = (json.choices && json.choices[0] && json.choices[0].message)
        ? json.choices[0].message.content
        : 'Khong nhan duoc phan hoi tu AI.';
      return { error: false, explanation: explanation };
    } else if (responseCode === 429) {
      return { error: true, message: 'He thong AI dang qua tai. Vui long thu lai sau vai giay.' };
    } else {
      return { error: true, message: 'Loi ket noi den may chu AI (ma loi: ' + responseCode + '). Vui long thu lai sau.' };
    }
  } catch (e) {
    return { error: true, message: 'Da xay ra loi khi goi AI. Vui long thu lai.' };
  }
}

// ---------- Utility: Include HTML parts ----------
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
