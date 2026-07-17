const CONFIG = {
  maxMessagesPerRequest: 100,
  defaultSenderName: '筑波大学アイドル研究会 Bombs!',
  defaultReplyTo: 'info@bombstsukuba.com',
};

function doGet() {
  return jsonResponse_({
    ok: true,
    name: 'Bombs benefit mail send API',
  });
}

function doPost(event) {
  try {
    const payload = parsePayload_(event);
    const messages = Array.isArray(payload.messages) ? payload.messages : [];

    if (messages.length === 0) {
      throw new Error('送信対象がありません。');
    }

    if (messages.length > CONFIG.maxMessagesPerRequest) {
      throw new Error(`1回の送信は${CONFIG.maxMessagesPerRequest}件までです。`);
    }

    const senderName = String(payload.senderName || CONFIG.defaultSenderName).trim();
    const replyTo = String(payload.replyTo || CONFIG.defaultReplyTo).trim();
    const results = messages.map((message) => sendOne_(message, senderName, replyTo));

    return jsonResponse_({
      ok: true,
      results,
    });
  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error.message,
    });
  }
}

function parsePayload_(event) {
  if (!event || !event.postData || !event.postData.contents) {
    throw new Error('リクエスト本文が空です。');
  }

  return JSON.parse(event.postData.contents);
}

function sendOne_(message, senderName, replyTo) {
  const rowNumber = Number(message.rowNumber || 0);

  try {
    const email = String(message.email || '').trim();
    const subject = String(message.subject || '').trim();
    const body = String(message.body || '').trim();

    if (!email) {
      throw new Error('メールアドレスが空です。');
    }

    if (!isValidEmail_(email)) {
      throw new Error(`メールアドレスの形式が不正です: ${email}`);
    }

    if (!subject) {
      throw new Error('件名が空です。');
    }

    if (!body) {
      throw new Error('本文が空です。');
    }

    GmailApp.sendEmail(email, subject, body, {
      name: senderName,
      replyTo,
    });

    return {
      ok: true,
      rowNumber,
      email,
    };
  } catch (error) {
    return {
      ok: false,
      rowNumber,
      email: String(message.email || '').trim(),
      error: error.message,
    };
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}
