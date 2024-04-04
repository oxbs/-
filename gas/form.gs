
//コンテナバインドスクリプト（フォームと連携）

//フォーム作成
function createForm() {
  form = FormApp.getActiveForm();
  form.setTitle('フォームタイトルを設定');
  form.setDescription(
    'フォームの補足情報を入力。'
  );
  sheet = fetchSheet('申込');
  questions = getQuestions(sheet);
  deleteQuestion(form);
  setQuestion(form, questions);
}

//シート情報取得
function fetchSheet(sheet_name) {
  spreadsheet = SpreadsheetApp.openById('1CbY18DhLDNLyDU13Ro5jVUeTITIOkhSqyE3-ZSV8y1A');
  sheet = spreadsheet.getSheetByName(sheet_name);
  return sheet;
}

//質問項目の取得
function getQuestions(sheet) {
  
  questions = [];

  //シート内容を行ごとに配列で取得
  question_values = sheet.getDataRange().getValues(); 
  
  //1行目を削除（1行目は項目名なので）
  question_values.shift(); 

  for (i = 0; i < question_values.length; i++) {
    questions[i] = [];
    questions[i]['title'] = question_values[i][0];
    questions[i]['type'] = question_values[i][1];

    //3列目以降を選択肢として取得
    choices = question_values[i].slice(2, question_values[i].length); 

    //isntBlank関数の条件（≠ブランク）を満たすデータを配列に格納
    questions[i]['choices'] = choices.filter(isntBlank); 
  }
  return questions;
}

function isntBlank(value) {
  return value != '';
}

//質問項目のフォーム設定
function setQuestion(form, questions) {
  questions.forEach(function (question) {
    if (question['type'] == '氏名') {
      form
        .addTextItem()
        .setTitle(question['title'])
        .setRequired(true);
    } else if (question['type'] == 'メールアドレス') {
      error_message = '入力されたメールアドレスは有効ではありません。';
      validation = FormApp.createTextValidation()
        .requireTextIsEmail()
        .setHelpText(error_message)
        .build();
      form
        .addTextItem()
        .setTitle(question['title'])
        .setRequired(true)
        .setValidation(validation);
    } else if (question['type'] == '備考') {
      form
        .addParagraphTextItem()
        .setTitle(question['title']);
    } else if (question['type'] == 'ラジオボタン') {
      form
        .addMultipleChoiceItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    } else if (question['type'] == 'プルダウン') {
      form
        .addListItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    } else if (question['type'] == 'チェックボックス') {
      form
        .addCheckboxItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    }
  });
}

//フォームの質問項目クリア
function deleteQuestion(form) {
  form_questions = form.getItems();
  form_questions.forEach(function (form_question) {
    form.deleteItem(form_question);
  });
}

//フォーム送信後の処理（トリガーでキャッチ・実行）
function reflectAnswer(e) {
  answers = e.response.getItemResponses();
  saveAnswers(answers);
}

//フォーム回答をスプレッドシートに反映
function saveAnswers(answers) {
  sheet = fetchSheet('回答');

  //回答シートの内容を行ごとに配列で取得
  answer_values = sheet.getDataRange().getValues(); 
  var i = answer_values.length +1;

  for (j = 1; j < 8; j++) {
    range = sheet.getRange(i,j);
    range.setValue(answers[j-1].getResponse()); 
  }

  //自動メール返信
  name = sheet.getRange(i,1).getValue(); //氏名を取得
  email = sheet.getRange(i,2).getValue(); //メールアドレスを取得
  level = sheet.getRange(i,5).getValue(); //受検級を取得

  //必要な情報を引数にしてメール処理へ
  postMail(name,email,level);

}

function postMail(name,email,level) {

  subject = "【自動返信】XXXXXXX";
  options = {
    name: "XXXXXX"
  };
  base_body = readBody();
  body = '';
  body = base_body.replace('[[名前]]', name);
  body = body.replace('[[受検級]]', level);
  tomail = email;
  GmailApp.sendEmail(tomail, subject, body, options);
}

//メール返信の本文をドキュメントファイルから取得
function readBody() {
  doc_url = 'https://docs.google.com/document/d/1MjhSqxTgs8AUj1BM6EQD9Q_6keeu_8q145qYG8pf_ic/edit';
  doc = DocumentApp.openByUrl(doc_url);
  return doc.getBody().getText();
}



