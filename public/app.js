const MAX_DAILY_SEND = 100;
const STORAGE_KEY = 'unidol-benefit-dispatch-console-gmail-api-experiment-v1';
const GMAIL_SEND_SCOPE = 'https://www.googleapis.com/auth/gmail.send';
const USERINFO_EMAIL_SCOPE = 'https://www.googleapis.com/auth/userinfo.email';

const state = {
  rows: [],
  activeId: null,
  sending: false,
  accessToken: '',
  tokenClient: null,
  signedInEmail: '',
};

const els = {
  sheetFile: document.getElementById('sheetFile'),
  oauthClientId: document.getElementById('oauthClientId'),
  authorizeButton: document.getElementById('authorizeButton'),
  senderName: document.getElementById('senderName'),
  replyTo: document.getElementById('replyTo'),
  subject: document.getElementById('subject'),
  openingText: document.getElementById('openingText'),
  sectionOneTitle: document.getElementById('sectionOneTitle'),
  sectionTwoTitle: document.getElementById('sectionTwoTitle'),
  signatureText: document.getElementById('signatureText'),
  totalCount: document.getElementById('totalCount'),
  selectedCount: document.getElementById('selectedCount'),
  errorCount: document.getElementById('errorCount'),
  recipientTableBody: document.getElementById('recipientTableBody'),
  previewTarget: document.getElementById('previewTarget'),
  previewSubject: document.getElementById('previewSubject'),
  previewBody: document.getElementById('previewBody'),
  authMessage: document.getElementById('authMessage'),
  limitMessage: document.getElementById('limitMessage'),
  sendButton: document.getElementById('sendButton'),
  sendDialog: document.getElementById('sendDialog'),
  sendDialogText: document.getElementById('sendDialogText'),
  confirmSendButton: document.getElementById('confirmSendButton'),
  saveTemplateButton: document.getElementById('saveTemplateButton'),
  selectValidButton: document.getElementById('selectValidButton'),
  clearSelectionButton: document.getElementById('clearSelectionButton'),
  downloadSampleButton: document.getElementById('downloadSampleButton'),
  toast: document.getElementById('toast'),
};

document.addEventListener('DOMContentLoaded', () => {
  loadSavedTemplate();
  bindEvents();
  render();
});

function bindEvents() {
  els.sheetFile.addEventListener('change', handleFileChange);
  els.saveTemplateButton.addEventListener('click', saveTemplate);
  els.selectValidButton.addEventListener('click', selectValidRows);
  els.clearSelectionButton.addEventListener('click', clearSelection);
  els.authorizeButton.addEventListener('click', authorizeGmail);
  els.sendButton.addEventListener('click', openSendDialog);
  els.confirmSendButton.addEventListener('click', sendSelectedRows);
  els.downloadSampleButton.addEventListener('click', downloadSampleCsv);

  [
    els.oauthClientId,
    els.subject,
    els.openingText,
    els.sectionOneTitle,
    els.sectionTwoTitle,
    els.signatureText,
  ].forEach((el) => el.addEventListener('input', () => {
    renderPreview();
    renderSendState();
  }));
}

async function handleFileChange(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  try {
    const rows = await readSheetFile(file);
    loadRows(rows);
    showToast(`${file.name} を読み込みました。`);
  } catch (error) {
    showToast(error.message);
  }
}

async function readSheetFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();

  if (extension === 'csv' || extension === 'tsv') {
    const text = await file.text();
    return parseDelimitedText(text, extension === 'tsv' ? '\t' : ',');
  }

  if (!window.XLSX) {
    throw new Error('Excel読込ライブラリの読み込みに失敗しました。CSVでアップロードしてください。');
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error('Excelファイルにシートがありません。');
  }

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    blankrows: false,
    raw: false,
  });
}

function loadRows(rawRows) {
  const normalized = rawRows
    .map((row) => Array.from({ length: 7 }, (_, index) => cleanCell(row[index])))
    .filter((row) => row.some(Boolean));

  if (normalized.length < 2) {
    throw new Error('2行目以降のデータがありません。');
  }

  const header = normalized[0];

  if (header[2]) {
    els.sectionOneTitle.value = header[2];
  }

  if (header[4]) {
    els.sectionTwoTitle.value = header[4];
  }

  state.rows = normalized.slice(1).map((row, index) => {
    const item = {
      id: `row-${index + 2}-${Date.now()}`,
      rowNumber: index + 2,
      email: row[0],
      recipientName: row[1],
      sectionOneName: row[2],
      sectionOneLink: row[3],
      sectionTwoName: row[4],
      sectionTwoLink: row[5],
      selected: parseCheckbox(row[6]),
      sent: false,
      error: '',
    };
    item.validationErrors = validateRow(item);
    return item;
  });

  state.activeId = state.rows[0]?.id ?? null;
  render();
}

function parseDelimitedText(text, delimiter) {
  const rows = [];
  let row = [];
  let value = '';
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const nextChar = text[index + 1];

    if (char === '"' && quoted && nextChar === '"') {
      value += '"';
      index += 1;
      continue;
    }

    if (char === '"') {
      quoted = !quoted;
      continue;
    }

    if (char === delimiter && !quoted) {
      row.push(value);
      value = '';
      continue;
    }

    if ((char === '\n' || char === '\r') && !quoted) {
      if (char === '\r' && nextChar === '\n') {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = '';
      continue;
    }

    value += char;
  }

  row.push(value);
  rows.push(row);
  return rows;
}

function render() {
  renderCounts();
  renderTable();
  renderPreview();
  renderSendState();
}

function renderCounts() {
  const selected = getSelectedRows();
  const errors = state.rows.filter((row) => row.validationErrors.length > 0);

  els.totalCount.textContent = state.rows.length;
  els.selectedCount.textContent = selected.length;
  els.errorCount.textContent = errors.length;
}

function renderTable() {
  if (state.rows.length === 0) {
    els.recipientTableBody.innerHTML = '<tr><td colspan="7" class="empty-state">ファイルをアップロードしてください。</td></tr>';
    return;
  }

  els.recipientTableBody.innerHTML = state.rows.map((row) => {
    const one = formatBenefitCell(row.sectionOneName, row.sectionOneLink);
    const two = formatBenefitCell(row.sectionTwoName, row.sectionTwoLink);
    const statuses = getRowStatuses(row);
    const needsReview = row.validationErrors.length > 0;

    return `
      <tr class="${row.id === state.activeId ? 'active' : ''} ${needsReview ? 'needs-review' : ''}" data-row-id="${row.id}">
        <td class="check-cell">
          <input type="checkbox" data-select-id="${row.id}" ${row.selected ? 'checked' : ''} ${row.validationErrors.length ? 'disabled' : ''}>
        </td>
        <td>${row.rowNumber}</td>
        <td>${escapeHtml(row.email || '-')}</td>
        <td>${escapeHtml(row.recipientName || '-')}</td>
        <td>${one}</td>
        <td>${two}</td>
        <td><div class="status-list">${statuses.map((status) => `<span class="status ${status.className}">${escapeHtml(status.label)}</span>`).join('')}</div></td>
      </tr>
    `;
  }).join('');

  els.recipientTableBody.querySelectorAll('tr[data-row-id]').forEach((tr) => {
    tr.addEventListener('click', (event) => {
      if (event.target.matches('input[type="checkbox"]')) {
        return;
      }
      state.activeId = tr.dataset.rowId;
      render();
    });
  });

  els.recipientTableBody.querySelectorAll('input[data-select-id]').forEach((checkbox) => {
    checkbox.addEventListener('change', () => {
      const row = state.rows.find((item) => item.id === checkbox.dataset.selectId);
      row.selected = checkbox.checked;
      renderCounts();
      renderSendState();
    });
  });
}

function renderPreview() {
  const row = state.rows.find((item) => item.id === state.activeId);

  els.previewSubject.textContent = els.subject.value || '-';

  if (!row) {
    els.previewTarget.textContent = '未選択';
    els.previewBody.textContent = '左の表から行を選択してください。';
    return;
  }

  els.previewTarget.textContent = `${row.rowNumber}行目`;
  els.previewBody.textContent = buildMessage(row).body;
}

function renderSendState() {
  const selected = getSelectedRows();
  const invalidSelected = selected.filter((row) => row.validationErrors.length > 0);
  const ready = selected.length > 0
    && selected.length <= MAX_DAILY_SEND
    && invalidSelected.length === 0
    && Boolean(state.accessToken)
    && Boolean(els.subject.value.trim())
    && !state.sending;

  if (state.signedInEmail) {
    els.authMessage.textContent = `${state.signedInEmail} で接続中です。`;
  } else {
    els.authMessage.textContent = 'Google OAuth クライアントIDを入力して接続してください。';
  }

  if (selected.length > MAX_DAILY_SEND) {
    els.limitMessage.textContent = `選択中: ${selected.length}件。100件以内に減らしてください。`;
  } else if (!state.accessToken) {
    els.limitMessage.textContent = 'Googleで接続してください。';
  } else if (!els.subject.value.trim()) {
    els.limitMessage.textContent = '件名を入力してください。';
  } else {
    els.limitMessage.textContent = `選択中: ${selected.length}件。100件まで送信できます。`;
  }

  els.sendButton.disabled = !ready;
}

function getRowStatuses(row) {
  if (row.sent) {
    return [{ label: '送信済み', className: 'ok' }];
  }

  if (row.error) {
    return [{ label: '送信エラー', className: 'error' }];
  }

  const statuses = [];

  row.validationErrors.forEach((label) => {
    statuses.push({ label, className: 'error' });
  });

  getBenefitWarnings(row).forEach((label) => {
    statuses.push({ label, className: 'warn' });
  });

  if (statuses.length > 0) {
    return statuses;
  }

  return [{ label: '送信可', className: 'ok' }];
}

function buildMessage(row) {
  const sections = [
    `${row.recipientName}様`,
    ...splitLines(els.openingText.value),
  ];

  appendLinkSection(sections, els.sectionOneTitle.value, row.sectionOneName, row.sectionOneLink);
  appendLinkSection(sections, els.sectionTwoTitle.value, row.sectionTwoName, row.sectionTwoLink);
  sections.push(...splitLines(els.signatureText.value));

  return {
    rowNumber: row.rowNumber,
    email: row.email,
    subject: els.subject.value.trim(),
    body: sections.join('\n'),
  };
}

function appendLinkSection(sections, title, subName, link) {
  if (!cleanCell(title) || !cleanCell(link)) {
    return;
  }

  sections.push(`[${cleanCell(title)}]`);

  if (cleanCell(subName)) {
    sections.push(cleanCell(subName));
  }

  sections.push(cleanCell(link));
}

function selectValidRows() {
  state.rows.forEach((row) => {
    row.selected = row.validationErrors.length === 0;
  });
  render();
}

function clearSelection() {
  state.rows.forEach((row) => {
    row.selected = false;
  });
  render();
}

function openSendDialog() {
  const selected = getSelectedRows();
  els.sendDialogText.textContent = `${selected.length}件のメールを送信します。実行しますか？`;
  els.sendDialog.showModal();
}

async function sendSelectedRows() {
  const messages = getSelectedRows().map(buildMessage);

  state.sending = true;
  renderSendState();

  try {
    const results = [];

    for (const message of messages) {
      results.push(await sendMessageWithGmailApi(message));
    }

    applySendResults(results);
    showToast('送信処理が完了しました。');
  } catch (error) {
    showToast(error.message);
  } finally {
    state.sending = false;
    render();
  }
}

function applySendResults(results) {
  results.forEach((result) => {
    const row = state.rows.find((item) => item.rowNumber === result.rowNumber);

    if (!row) {
      return;
    }

    row.sent = result.ok === true;
    row.error = result.ok ? '' : result.error || '送信に失敗しました。';
    row.selected = !row.sent;
  });
}

function getSelectedRows() {
  return state.rows.filter((row) => row.selected);
}

function validateRow(row) {
  const errors = [];

  if (!row.email) {
    errors.push('メールなし');
  }

  if (!row.recipientName) {
    errors.push('宛名なし');
  }

  return errors;
}

function getBenefitWarnings(row) {
  const warnings = [];

  appendBenefitWarning(warnings, '1つ目', row.sectionOneName, row.sectionOneLink);
  appendBenefitWarning(warnings, '2つ目', row.sectionTwoName, row.sectionTwoLink);

  return warnings;
}

function appendBenefitWarning(warnings, label, subName, link) {
  if (!cleanCell(link)) {
    warnings.push(`${label}リンクなし`);
  }

  if (!cleanCell(subName)) {
    warnings.push(`${label}サブ名なし`);
  }
}

function saveTemplate() {
  const data = {
    oauthClientId: els.oauthClientId.value,
    senderName: els.senderName.value,
    replyTo: els.replyTo.value,
    subject: els.subject.value,
    openingText: els.openingText.value,
    sectionOneTitle: els.sectionOneTitle.value,
    sectionTwoTitle: els.sectionTwoTitle.value,
    signatureText: els.signatureText.value,
  };

  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  showToast('設定を保存しました。');
}

function loadSavedTemplate() {
  const saved = localStorage.getItem(STORAGE_KEY);

  if (!saved) {
    return;
  }

  try {
    const data = JSON.parse(saved);
    Object.entries(data).forEach(([key, value]) => {
      if (els[key] && typeof value === 'string') {
        els[key].value = value;
      }
    });
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function authorizeGmail() {
  const clientId = els.oauthClientId.value.trim();

  if (!clientId) {
    showToast('Google OAuth クライアントIDを入力してください。');
    return;
  }

  if (!window.google || !google.accounts || !google.accounts.oauth2) {
    showToast('Google Identity Servicesを読み込めませんでした。');
    return;
  }

  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: `${GMAIL_SEND_SCOPE} ${USERINFO_EMAIL_SCOPE}`,
    callback: async (response) => {
      if (response.error) {
        showToast(response.error);
        return;
      }

      state.accessToken = response.access_token;
      await loadSignedInEmail();
      renderSendState();
    },
  });

  state.tokenClient.requestAccessToken({ prompt: 'consent' });
}

async function loadSignedInEmail() {
  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: {
        Authorization: `Bearer ${state.accessToken}`,
      },
    });

    if (!response.ok) {
      return;
    }

    const profile = await response.json();
    state.signedInEmail = profile.email || '';
  } catch {
    state.signedInEmail = '';
  }
}

async function sendMessageWithGmailApi(message) {
  try {
    const raw = buildRawEmail(message);
    const response = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ raw }),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error?.message || 'Gmail APIで送信に失敗しました。');
    }

    return {
      ok: true,
      rowNumber: message.rowNumber,
      email: message.email,
    };
  } catch (error) {
    return {
      ok: false,
      rowNumber: message.rowNumber,
      email: message.email,
      error: error.message,
    };
  }
}

function buildRawEmail(message) {
  const headers = [
    ['To', message.email],
    ['Subject', message.subject],
    ['MIME-Version', '1.0'],
    ['Content-Type', 'text/plain; charset=UTF-8'],
    ['Content-Transfer-Encoding', '8bit'],
  ];
  const senderName = els.senderName.value.trim();
  const replyTo = els.replyTo.value.trim();

  if (state.signedInEmail) {
    headers.unshift(['From', formatAddress(senderName, state.signedInEmail)]);
  }

  if (replyTo) {
    headers.push(['Reply-To', replyTo]);
  }

  return base64UrlEncode(`${headers.map(([key, value]) => `${key}: ${encodeHeaderValue(value)}`).join('\r\n')}\r\n\r\n${message.body}`);
}

function formatAddress(name, email) {
  if (!name) {
    return email;
  }

  return `${encodeMimeWord(name)} <${email}>`;
}

function encodeHeaderValue(value) {
  return /[^\x20-\x7e]/.test(value) ? encodeMimeWord(value) : value;
}

function encodeMimeWord(value) {
  return `=?UTF-8?B?${base64EncodeUtf8(value)}?=`;
}

function base64EncodeUtf8(value) {
  const bytes = new TextEncoder().encode(value);
  let binary = '';

  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary);
}

function base64UrlEncode(value) {
  return base64EncodeUtf8(value)
    .replaceAll('+', '-')
    .replaceAll('/', '_')
    .replaceAll('=', '');
}

function downloadSampleCsv() {
  const csv = [
    ['メールアドレス', '宛名', 'メッセージ動画', '特典リンク', '待ち受け画像', '特典リンク', '送信対象'],
    ['sample@example.com', '山田', 'あやか', 'https://example.com/video', 'りな', 'https://example.com/wallpaper', 'TRUE'],
    ['sample2@example.com', '佐藤', '', 'https://example.com/video2', '', 'https://example.com/wallpaper2', 'FALSE'],
    ['sample3@example.com', '鈴木', 'みほ', '', 'りな', '', 'FALSE'],
  ].map((row) => row.map(csvEscape).join(',')).join('\n');

  const blob = new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = 'benefit-mail-sample.csv';
  anchor.click();
  URL.revokeObjectURL(url);
}

function splitLines(value) {
  return value.split(/\r?\n/).map(cleanCell).filter(Boolean);
}

function parseCheckbox(value) {
  return ['true', '1', 'yes', 'y', 'checked', '送信'].includes(cleanCell(value).toLowerCase());
}

function cleanCell(value) {
  return String(value ?? '').trim();
}

function formatBenefitCell(subName, link) {
  const lines = [];

  if (cleanCell(subName)) {
    lines.push(escapeHtml(subName));
  }

  if (cleanCell(link)) {
    lines.push(escapeHtml(link));
  }

  return lines.length > 0 ? lines.join('<br>') : '-';
}

function csvEscape(value) {
  return `"${String(value).replaceAll('"', '""')}"`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add('visible');
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    els.toast.classList.remove('visible');
  }, 3600);
}
