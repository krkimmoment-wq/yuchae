/**
 * 코치진 단톡방 셀프체크 — Google Sheets 동기화 백엔드
 *
 * 사용법:
 *  1) Google Drive에서 새 Google Sheets 생성
 *  2) 확장 프로그램(Extensions) > Apps Script 클릭
 *  3) 기본 코드 전부 지우고 이 파일 내용 통째로 붙여넣기
 *  4) 저장 (디스크 아이콘 또는 Ctrl+S)
 *  5) 우측 상단 [배포(Deploy)] > [새 배포(New deployment)]
 *  6) 톱니바퀴 > [웹 앱(Web app)] 선택
 *  7) 설정:
 *      - 설명: 아무거나
 *      - 다음 사용자 권한으로 실행: 나
 *      - 액세스 권한이 있는 사용자: '모든 사용자(액세스 권한 없는 사용자 포함)' (또는 본인 도메인)
 *  8) [배포] 클릭 > 권한 승인
 *  9) 발급된 웹 앱 URL(/exec 로 끝남)을
 *     group_chat_self_check.html 의 var SYNC_URL = '' 안에 붙여넣기
 *
 * 코드 수정 후 재배포할 때:
 *  - [배포] > [배포 관리] > 연필 아이콘 > 버전: 새 버전 > [배포]
 *  - URL은 그대로 유지됨 (절대 변하지 않음)
 */

var SHEET_NAME = 'checks';
var META_SHEET_NAME = 'meta';
var COLUMNS = [
  'id', 'date', 'coach', 'level', 'cohort', 'groupDetail', 'topic',
  'uploaded', 'uploadTime', 'reaction', 'reactionCount', 'reason',
  'selfCheck', 'leaderConfirm'
];

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    var sh = getSheet_();
    var lastRow = sh.getLastRow();
    var data = [];
    if (lastRow >= 2) {
      var values = sh.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
      data = values.map(function(row) {
        var obj = {};
        COLUMNS.forEach(function(col, i) {
          var v = row[i];
          if (col === 'uploaded' || col === 'selfCheck' || col === 'leaderConfirm') {
            obj[col] = (v === true || v === 'TRUE' || v === 'true');
          } else if (col === 'id' || col === 'reactionCount') {
            obj[col] = (v === '' || v === null) ? 0 : Number(v);
          } else {
            obj[col] = (v == null) ? '' : String(v);
          }
        });
        return obj;
      }).filter(function(r) { return r.id; });
    }
    return jsonOut_({
      ok: true,
      data: data,
      updatedAt: getMeta_('updatedAt') || ''
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (!Array.isArray(body.data)) {
      return jsonOut_({ ok: false, error: 'data must be an array' });
    }

    var sh = getSheet_();

    // 헤더 보장
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);

    // 기존 데이터 영역 클리어
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      sh.getRange(2, 1, lastRow - 1, COLUMNS.length).clearContent();
    }

    // 새 데이터 쓰기
    if (body.data.length > 0) {
      var rows = body.data.map(function(item) {
        return COLUMNS.map(function(col) {
          var v = item[col];
          return (v === undefined || v === null) ? '' : v;
        });
      });
      sh.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
    }

    var now = new Date().toISOString();
    setMeta_('updatedAt', now);

    return jsonOut_({ ok: true, updatedAt: now, count: body.data.length });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getMetaSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(META_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(META_SHEET_NAME);
    sh.hideSheet();
  }
  return sh;
}

function getMeta_(key) {
  var sh = getMetaSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

function setMeta_(key, value) {
  var sh = getMetaSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
