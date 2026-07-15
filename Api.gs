function doGet(e) {
  try {
    const action = getAction_(e);

    switch (action) {
      case 'ping':
        return jsonOutput_({ ok: true, message: 'alive' });

      case 'master':
        return jsonOutput_(getMasterData());

      case 'total':
        return jsonOutput_(
          getDayTotal(
            getParam_(e, 'token'),
            getParam_(e, 'workDate')
          )
        );

      case 'historyDay':
        return jsonOutput_(
          getHistoryDay(
            getParam_(e, 'token'),
            getParam_(e, 'workDate')
          )
        );

      case 'historyWeek':
        return jsonOutput_(
          getHistoryWeek(
            getParam_(e, 'token'),
            getParam_(e, 'baseDate')
          )
        );

      case 'requestList':
        return jsonOutput_(getRequestMasterList(getParam_(e, 'token')));

      default:
        return jsonOutput_({
          ok: false,
          message: '不明なGETアクションです。',
        });
    }
  } catch (e) {
    console.error(e);
    return jsonOutput_({
      ok: false,
      message: e.message || 'GET処理でエラーが発生しました。',
    });
  }
}

function doPost(e) {
  try {
    const action = getAction_(e);
    const body = parseRequestBody_(e);

    switch (action) {
      case 'login':
        return jsonOutput_(verifyLogin(body.name));

      case 'saveReport':
        return jsonOutput_(saveReport(body));

      case 'requestAdd':
        return jsonOutput_(addRequestNo(body.token, body.requestNo));

      case 'requestDisable':
        return jsonOutput_(disableRequestNo(body.token, body.requestNo));

      default:
        return jsonOutput_({
          ok: false,
          message: '不明なPOSTアクションです。',
        });
    }
  } catch (e) {
    console.error(e);
    return jsonOutput_({
      ok: false,
      message: e.message || 'POST処理でエラーが発生しました。',
    });
  }
}

function getAction_(e) {
  return String((e && e.parameter && e.parameter.action) || '').trim();
}

function getParam_(e, key) {
  return String((e && e.parameter && e.parameter[key]) || '').trim();
}

function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  const raw = String(e.postData.contents || '').trim();
  if (!raw) {
    return {};
  }

  return JSON.parse(raw);
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
