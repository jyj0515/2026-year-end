/**
 * 올해도 거친 세상속에서 살아남았다 — 점검 메일 (구글 시트 + Apps Script)
 *
 * 하는 일
 *  - doPost   : 웹앱에서 보낸 메일 신청을 시트에 저장 (같은 이메일이면 최신 내용으로 덮어씀)
 *  - doGet    : 메일 속 '수신 거부' 링크 처리 (해당 행 삭제)
 *  - sendJune / sendDecember : 2027-06-01, 2027-12-01 09:00(KST)에 자동 발송
 *  - 12월 발송이 끝나면 시트의 모든 신청 정보를 삭제 (기획서 6장 보관 기간)
 *
 * 처음 한 번: setup() 실행 → 웹앱으로 배포 → 배포 URL을 index.html 의 SUBSCRIBE_URL 에 붙여넣기
 * 자세한 순서는 SETUP_EMAIL.md 참고
 */

// ===== 설정 =====
const SERVICE_NAME = '올해도 거친 세상속에서 살아남았다';

// 사이트를 웹에 올린 주소 (예: 'https://내아이디.github.io/2026_v/')
// 메일에 들어갈 링크가 이 주소로 시작할 때만 신청을 받아요. (다른 사이트 링크를 넣은 스팸 방지)
const SITE_URL = 'https://jyj0515.github.io/2026-year-end/';

const SEND_JUNE = new Date('2027-06-01T09:00:00+09:00');
const SEND_DEC = new Date('2027-12-01T09:00:00+09:00');

// ===== 내부 상수 =====
const SHEET_NAME = 'subscribers';
const HEADERS = ['createdAt', 'updatedAt', 'email', 'nextWord', 'goal1', 'goal2', 'goal3', 'link', 'token', 'status', 'junSentAt', 'decSentAt'];
const COL = { createdAt: 1, updatedAt: 2, email: 3, nextWord: 4, goal1: 5, link: 8, token: 9, status: 10, junSentAt: 11, decSentAt: 12 };
const LIMITS = { nextWord: 40, goal: 25, link: 4000, email: 254 };

/** 처음 한 번 실행: 시트 준비 + 6월/12월 발송 예약 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('구글 시트의 [확장 프로그램 > Apps Script]에서 만든 스크립트에서 실행해 주세요.');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  getSheet_();
  resetTrigger_('sendJune', SEND_JUNE);
  resetTrigger_('sendDecember', SEND_DEC);
  Logger.log('준비 완료. 6월 발송: %s / 12월 발송: %s', SEND_JUNE, SEND_DEC);
}

/** 내 메일함으로 6월 메일 미리보기를 보내요 (시트 첫 신청자 내용 또는 예시) */
function sendTestToMe() {
  const me = Session.getEffectiveUser().getEmail();
  const sh = getSheet_();
  const row = sh.getLastRow() > 1
    ? sh.getRange(2, 1, 1, HEADERS.length).getValues()[0]
    : ['', '', me, '걱정은 조금 덜고, 웃으며 긍정적으로!', '일주일에 2번, 30분씩 산책하기', '새로운 취미 하나 시작하기', '한 달에 5만 원씩 저축하기', SITE_URL || 'https://example.com/', 'test-token', 'active', '', ''];
  MailApp.sendEmail({ to: me, subject: '[미리보기] ' + subject_('jun'), htmlBody: mailHtml_('jun', row), name: SERVICE_NAME });
  MailApp.sendEmail({ to: me, subject: '[미리보기] ' + subject_('dec'), htmlBody: mailHtml_('dec', row), name: SERVICE_NAME });
  Logger.log('%s 로 미리보기 2통을 보냈어요.', me);
}

// ===== 웹앱: 신청 받기 =====
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.website) return json_({ ok: true }); // 봇 차단용 숨은 칸

    const email = String(body.email || '').trim().toLowerCase();
    if (email.length > LIMITS.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'invalid_email' });
    if (body.consent !== true) return json_({ ok: false, error: 'consent_required' });

    const goals = (Array.isArray(body.goals) ? body.goals : []).slice(0, 3).map(function (g) { return clean_(g, LIMITS.goal); });
    while (goals.length < 3) goals.push('');
    if (!goals.some(Boolean)) return json_({ ok: false, error: 'no_goals' });
    const nextWord = clean_(body.nextWord, LIMITS.nextWord);

    const link = String(body.link || '');
    if (!SITE_URL) return json_({ ok: false, error: 'site_url_not_set' });
    if (link.length > LIMITS.link || link.indexOf(SITE_URL) !== 0) return json_({ ok: false, error: 'invalid_link' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet_();
      const now = new Date();
      const emails = sh.getLastRow() > 1 ? sh.getRange(2, COL.email, sh.getLastRow() - 1, 1).getValues() : [];
      let row = -1;
      for (let i = 0; i < emails.length; i++) if (emails[i][0] === email) { row = i + 2; break; }
      if (row > 0) {
        sh.getRange(row, COL.updatedAt, 1, 6).setValues([[now, email, nextWord, goals[0], goals[1], goals[2]]]);
        sh.getRange(row, COL.link).setValue(link);
        sh.getRange(row, COL.status).setValue('active');
      } else {
        sh.appendRow([now, now, email, nextWord, goals[0], goals[1], goals[2], link, Utilities.getUuid(), 'active', '', '']);
      }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

// ===== 웹앱: 수신 거부 =====
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'unsubscribe' && p.token) {
    const sh = getSheet_();
    const last = sh.getLastRow();
    const tokens = last > 1 ? sh.getRange(2, COL.token, last - 1, 1).getValues() : [];
    let found = false;
    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i][0] === p.token) { sh.deleteRow(i + 2); found = true; break; }
    }
    return page_(found ? '수신을 거부했어요' : '이미 처리됐어요',
      found ? '등록된 이메일과 목표를 삭제했어요. 더 이상 메일이 가지 않아요.' : '이미 수신 거부했거나 삭제된 정보예요.');
  }
  return page_(SERVICE_NAME, '점검 메일 서버가 동작 중이에요.');
}

// ===== 자동 발송 =====
function sendJune() { sendBatch_('jun'); }
function sendDecember() { sendBatch_('dec'); }

function sendBatch_(kind) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const sentCol = kind === 'jun' ? COL.junSentAt : COL.decSentAt;
  let left = 0;
  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (r[COL.status - 1] !== 'active' || r[sentCol - 1]) continue;
      if (MailApp.getRemainingDailyQuota() < 1) { left++; continue; }
      try {
        MailApp.sendEmail({ to: r[COL.email - 1], subject: subject_(kind), htmlBody: mailHtml_(kind, r), name: SERVICE_NAME });
        sh.getRange(i + 2, sentCol).setValue(new Date());
      } catch (err) {
        console.error(err);
        left++;
      }
    }
  }
  const handler = kind === 'jun' ? 'sendJune' : 'sendDecember';
  if (left > 0) {
    // 하루 발송 한도에 걸린 경우: 내일 같은 시각에 남은 사람에게 이어서 발송
    resetTrigger_(handler, new Date(Date.now() + 24 * 60 * 60 * 1000));
  } else {
    deleteTriggers_(handler);
    if (kind === 'dec') purgeAll_(); // 12월 발송 완료 → 모든 신청 정보 삭제
  }
}

function purgeAll_() {
  const sh = getSheet_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  Logger.log('12월 발송 완료: 신청 정보를 모두 삭제했어요.');
}

// ===== 메일 내용 =====
function subject_(kind) {
  return kind === 'jun'
    ? '[' + SERVICE_NAME + '] 6월, 목표 잘 지키고 있나요?'
    : '[' + SERVICE_NAME + '] 2027년, 얼마나 해냈는지 돌아봐요';
}

function mailHtml_(kind, r) {
  const nextWord = r[COL.nextWord - 1];
  const goals = [r[COL.goal1 - 1], r[COL.goal1], r[COL.goal1 + 1]].filter(String);
  let link = String(r[COL.link - 1] || SITE_URL);
  if (kind === 'dec') link = link.replace('#june&', '#'); // 12월은 결과 화면으로
  const unsub = ScriptApp.getService().getUrl() + '?action=unsubscribe&token=' + encodeURIComponent(r[COL.token - 1]);

  const intro = kind === 'jun'
    ? '반년 전, 2026년의 내가 2027년의 나에게 남긴 약속이에요.<br>지금까지 잘 지키고 있나요? 다시 시작해도 괜찮아요.'
    : '2027년도 거의 다 지나갔어요.<br>2026년의 내가 남긴 약속, 얼마나 해냈는지 돌아봐요.';
  const btn = kind === 'jun' ? '지금 돌아보기' : '내 기록 보러 가기';

  const goalItems = goals.map(function (g, i) {
    return '<tr><td style="padding:8px 0;vertical-align:top;width:28px"><span style="display:inline-block;width:22px;height:22px;border-radius:11px;background:#CB513B;color:#fff;font-size:12px;font-weight:700;text-align:center;line-height:22px">' + (i + 1) + '</span></td>'
      + '<td style="padding:8px 0;font-size:15px;color:#18181b">' + esc_(g) + '</td></tr>';
  }).join('');

  return ''
    + '<div style="background:#E9E1CE;padding:32px 16px;font-family:-apple-system,BlinkMacSystemFont,\'Apple SD Gothic Neo\',\'Noto Sans KR\',sans-serif">'
    + '<div style="max-width:440px;margin:0 auto;background:#ffffff;border-radius:16px;padding:28px 24px">'
    + '<p style="margin:0 0 6px;font-size:13px;color:#6E6656">' + esc_(SERVICE_NAME) + '</p>'
    + '<h1 style="margin:0 0 12px;font-size:22px;line-height:1.35;color:#18181b">' + (kind === 'jun' ? '지금까지 잘 해내고 있나요?' : '올해도 내가 해냈나요?') + '</h1>'
    + '<p style="margin:0 0 20px;font-size:14px;line-height:1.7;color:#6E6656">' + intro + '</p>'
    + (nextWord ? '<p style="margin:0 0 4px;font-size:12px;color:#6E6656">2027년의 나에게</p><p style="margin:0 0 18px;font-size:16px;font-weight:600;color:#18181b">' + esc_(nextWord) + '</p>' : '')
    + '<p style="margin:0 0 4px;font-size:12px;color:#6E6656">2027년에 해보고 싶은 3가지</p>'
    + '<table role="presentation" style="border-collapse:collapse;margin:0 0 24px">' + goalItems + '</table>'
    + '<a href="' + esc_(link) + '" style="display:block;text-align:center;background:#CB513B;color:#ffffff;text-decoration:none;font-weight:700;font-size:15px;padding:14px;border-radius:10px">' + btn + '</a>'
    + '</div>'
    + '<p style="max-width:440px;margin:16px auto 0;font-size:12px;line-height:1.6;color:#8C836F;text-align:center">'
    + '이 메일은 신청하신 분께만 2027년 6월 1일, 12월 1일 두 번 보내져요.<br>12월 메일을 보낸 뒤 이메일과 목표 정보는 삭제돼요.<br>'
    + '<a href="' + esc_(unsub) + '" style="color:#8C836F">수신 거부하고 정보 삭제하기</a></p>'
    + '</div>';
}

// ===== 도우미 =====
function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function resetTrigger_(handler, when) {
  deleteTriggers_(handler);
  ScriptApp.newTrigger(handler).timeBased().at(when).create();
}

function deleteTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

// 시트에 넣는 글: 길이 제한 + 수식으로 해석되지 않게
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/^'/, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function page_(title, msg) {
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<div style="font-family:-apple-system,sans-serif;max-width:420px;margin:15vh auto;padding:24px;text-align:center">'
    + '<h2 style="margin:0 0 8px">' + esc_(title) + '</h2><p style="color:#6E6656">' + esc_(msg) + '</p></div>'
  ).setTitle(title);
}
