const CONFIG = {
  SPREADSHEET_ID: "1O60-AFjH0rAsDWk2QRwC-HuGKByQObLWmV-b-WhUQVA",
  SHEET_NAME: "代表信息汇总",
  RECIPIENT_EMAIL: "info@acbec.com.au",
  EMAIL_SUBJECT: "ACBEC 理事单位代表信息提交"
};

const FIELDS = [
  { key: "中文名字 / Chinese Name", header: "中文名字" },
  { key: "英文名字 / English Name", header: "英文名字" },
  { key: "出生地 / Place of Birth", header: "出生地" },
  { key: "生日 / Date of Birth", header: "生日" },
  { key: "祖籍地 / Ancestral Home", header: "祖籍地" },
  { key: "生活/居住地 / Residence", header: "生活/居住地" },
  { key: "国籍 / Nationality", header: "国籍" },
  { key: "目前身份 / Current Status", header: "目前身份" },
  { key: "护照号 / Passport No.", header: "护照号" },
  { key: "大陆身份证 / Mainland China ID", header: "大陆身份证" },
  { key: "外国手机号 / Overseas Mobile", header: "外国手机号" },
  { key: "中国手机号 / China Mobile", header: "中国手机号" },
  { key: "email", header: "Email" },
  { key: "微信号 / WeChat ID", header: "微信号" },
  { key: "企业职务 / Corporate Role", header: "企业职务" },
  { key: "社会职务 / Social Role", header: "社会职务" },
  { key: "荣誉头衔 / Honorary Titles", header: "荣誉头衔" },
  { key: "隐私确认 / Privacy Consent", header: "隐私确认" }
];

const EXTRA_HEADERS = ["提交时间", "提交来源页面"];

function doGet() {
  return jsonResponse({
    ok: true,
    service: "ACBEC representative form backend"
  });
}

function doPost(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};

    if (String(params._honey || "").trim()) {
      return jsonResponse({ ok: true, skipped: true });
    }

    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    const submittedAt = new Date();
    const row = [
      submittedAt,
      ...FIELDS.map((field) => cleanValue(params[field.key])),
      cleanValue(params._page_url)
    ];

    sheet.appendRow(row);
    autoResize(sheet);
    sendNotificationEmail(params, submittedAt);

    return jsonResponse({ ok: true });
  } catch (error) {
    console.error(error);
    return jsonResponse({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return spreadsheet.getSheetByName(CONFIG.SHEET_NAME) ||
    spreadsheet.insertSheet(CONFIG.SHEET_NAME);
}

function ensureHeaders(sheet) {
  const headers = [EXTRA_HEADERS[0], ...FIELDS.map((field) => field.header), EXTRA_HEADERS[1]];
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = current.every((cell) => String(cell || "").trim() === "");

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#eaf2fb");
  }
}

function autoResize(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 0) {
    sheet.autoResizeColumns(1, lastColumn);
  }
}

function sendNotificationEmail(params, submittedAt) {
  const subjectName = cleanValue(params["中文名字 / Chinese Name"]) ||
    cleanValue(params["英文名字 / English Name"]) ||
    "新提交";

  const htmlRows = [
    ["提交时间", Utilities.formatDate(submittedAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")],
    ...FIELDS.map((field) => [field.header, cleanValue(params[field.key])]),
    ["提交来源页面", cleanValue(params._page_url)]
  ].map(([label, value]) => `
    <tr>
      <th style="text-align:left;padding:8px;border:1px solid #d9e2ef;background:#f4f7fb;">${escapeHtml(label)}</th>
      <td style="padding:8px;border:1px solid #d9e2ef;">${escapeHtml(value)}</td>
    </tr>
  `).join("");

  MailApp.sendEmail({
    to: CONFIG.RECIPIENT_EMAIL,
    subject: `${CONFIG.EMAIL_SUBJECT} - ${subjectName}`,
    replyTo: cleanValue(params.email) || undefined,
    name: "ACBEC Representative Form",
    htmlBody: `
      <p>收到一份新的 ACBEC 理事单位代表信息提交，内容已同步写入 Google Sheet。</p>
      <table style="border-collapse:collapse;font-family:Arial,'Microsoft YaHei',sans-serif;font-size:14px;">
        ${htmlRows}
      </table>
    `
  });
}

function cleanValue(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return cleanValue(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
