const CONFIG = {
  subject: 'UNIDOLデジタル特典送付のお知らせ',
  senderName: 'UNIDOLデジタル特典送付システム',
  replyTo: '',
  columns: {
    email: 1,
    name: 2,
    videoTitle: 3,
    videoName: 3,
    videoLink: 4,
    wallpaperTitle: 5,
    wallpaperName: 5,
    wallpaperLink: 6,
    selected: 7,
    status: 8,
    sentAt: 9,
    error: 10,
  },
  statusSent: '送信済み',
  statusError: 'エラー',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('特典メール')
    .addItem('チェックボックスを準備', 'prepareRecipientCheckboxes')
    .addItem('選択行をプレビュー', 'previewSelectedRow')
    .addItem('チェック済みの未送信者に送信', 'sendBenefitEmails')
    .addToUi();
}

function prepareRecipientCheckboxes() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  ensureControlHeaders_(sheet);

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('チェックボックスを追加する行がありません。');
    return;
  }

  const checkboxRange = sheet.getRange(2, CONFIG.columns.selected, lastRow - 1, 1);
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  const values = checkboxRange.getValues().map(([value]) => [
    value === true || value === false ? value : false,
  ]);

  checkboxRange.setDataValidation(checkboxRule);
  checkboxRange.setValues(values);

  SpreadsheetApp.getUi().alert('G列に送信対象のチェックボックスを準備しました。');
}

function previewSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('2行目以降の送信対象行を選択してください。');
    return;
  }

  const message = buildMessageForRow_(sheet, row);

  SpreadsheetApp.getUi().alert(
    `宛先: ${message.email}\n件名: ${CONFIG.subject}\n\n${message.body}`,
  );
}

function sendBenefitEmails() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'G列にチェックがあり、H列が「送信済み」ではない行にメールを送信します。実行しますか？',
    ui.ButtonSet.YES_NO,
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('送信対象の行がありません。');
    return;
  }

  ensureControlHeaders_(sheet);

  let sentCount = 0;
  let skippedCount = 0;
  let uncheckedCount = 0;
  let errorCount = 0;

  for (let row = 2; row <= lastRow; row += 1) {
    const selected = sheet.getRange(row, CONFIG.columns.selected).getValue() === true;
    const status = getCellText_(sheet, row, CONFIG.columns.status);

    if (!selected) {
      uncheckedCount += 1;
      continue;
    }

    if (status === CONFIG.statusSent) {
      skippedCount += 1;
      continue;
    }

    try {
      const message = buildMessageForRow_(sheet, row);

      GmailApp.sendEmail(message.email, CONFIG.subject, message.body, buildSendOptions_());

      sheet.getRange(row, CONFIG.columns.status).setValue(CONFIG.statusSent);
      sheet.getRange(row, CONFIG.columns.sentAt).setValue(new Date());
      sheet.getRange(row, CONFIG.columns.error).clearContent();
      sentCount += 1;
    } catch (error) {
      sheet.getRange(row, CONFIG.columns.status).setValue(CONFIG.statusError);
      sheet.getRange(row, CONFIG.columns.error).setValue(error.message);
      errorCount += 1;
    }
  }

  ui.alert(
    `送信完了\n送信: ${sentCount}件\n未チェック: ${uncheckedCount}件\n送信済みスキップ: ${skippedCount}件\nエラー: ${errorCount}件`,
  );
}

function buildMessageForRow_(sheet, row) {
  const email = getCellText_(sheet, row, CONFIG.columns.email);
  const recipientName = getCellText_(sheet, row, CONFIG.columns.name);
  const videoTitle = getCellText_(sheet, 1, CONFIG.columns.videoTitle);
  const videoName = getCellText_(sheet, row, CONFIG.columns.videoName);
  const videoLink = getCellText_(sheet, row, CONFIG.columns.videoLink);
  const wallpaperTitle = getCellText_(sheet, 1, CONFIG.columns.wallpaperTitle);
  const wallpaperName = getCellText_(sheet, row, CONFIG.columns.wallpaperName);
  const wallpaperLink = getCellText_(sheet, row, CONFIG.columns.wallpaperLink);

  if (!email) {
    throw new Error('A列のメールアドレスが空です。');
  }

  if (!isValidEmail_(email)) {
    throw new Error(`メールアドレスの形式が不正です: ${email}`);
  }

  if (!recipientName) {
    throw new Error('B列の名前が空です。');
  }

  const sections = [
    `${recipientName}様`,
    'この度は応援いただき、誠にありがとうございました。',
    'デジタル特典を本メール下部に記載しておりますので、ぜひご覧ください。',
    '今後とも温かいご声援を何卒よろしくお願い申し上げます。',
  ];

  appendLinkSection_(sections, videoTitle, videoName, videoLink);
  appendLinkSection_(sections, wallpaperTitle, wallpaperName, wallpaperLink);

  sections.push('送信チーム名', '連絡先メールアドレス');

  return {
    email,
    body: sections.join('\n'),
  };
}

function buildSendOptions_() {
  const options = {
    name: CONFIG.senderName,
  };

  if (CONFIG.replyTo) {
    options.replyTo = CONFIG.replyTo;
  }

  return options;
}

function appendLinkSection_(sections, title, subName, link) {
  if (!title || !link) {
    return;
  }

  sections.push(`[${title}]`);

  if (subName) {
    sections.push(subName);
  }

  sections.push(link);
}

function ensureControlHeaders_(sheet) {
  sheet.getRange(1, CONFIG.columns.selected).setValue('送信対象');
  sheet.getRange(1, CONFIG.columns.status).setValue('送信ステータス');
  sheet.getRange(1, CONFIG.columns.sentAt).setValue('送信日時');
  sheet.getRange(1, CONFIG.columns.error).setValue('エラー内容');
}

function getCellText_(sheet, row, column) {
  return String(sheet.getRange(row, column).getDisplayValue()).trim();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}
