/**
 * P523 Drive API v2 compatibility helpers.
 *
 * appsscript.json currently enables Advanced Drive API v2.  Several legacy
 * UrlFetch call sites were written in Drive API v3 style (files/name/webViewLink).
 * These helpers keep the existing call-site object shape stable while sending
 * v2-compatible endpoints, fields, payload keys and query keys.
 */
function driveV2CompatBuildUrl_(path, upload) {
  return 'https://www.googleapis.com/' + (upload ? 'upload/' : '') + 'drive/v2/' + driveV2CompatPath_(path || '');
}

function driveV2CompatPath_(path) {
  var raw = String(path || '');
  var queryIndex = raw.indexOf('?');
  var resourcePath = queryIndex >= 0 ? raw.substring(0, queryIndex) : raw;
  var queryString = queryIndex >= 0 ? raw.substring(queryIndex + 1) : '';
  var isTeamDriveList = /^drives\b/.test(resourcePath);

  if (isTeamDriveList) {
    resourcePath = resourcePath.replace(/^drives\b/, 'teamdrives');
  }

  resourcePath = resourcePath
    .replace(/upload\/drive\/v3/g, 'upload/drive/v2')
    .replace(/drive\/v3/g, 'drive/v2');

  if (!queryString) return resourcePath;

  var translatedParams = queryString.split('&').map(function(part) {
    if (!part) return part;

    var separatorIndex = part.indexOf('=');
    if (separatorIndex < 0) return part;

    var rawKey = part.substring(0, separatorIndex);
    var rawValue = part.substring(separatorIndex + 1);
    var decodedKey = driveV2CompatSafeDecode_(rawKey);
    var decodedValue = driveV2CompatSafeDecode_(rawValue);

    if (decodedKey === 'q') {
      decodedValue = driveV2CompatTranslateQueryExpression_(decodedValue, isTeamDriveList);
      return rawKey + '=' + encodeURIComponent(decodedValue);
    }

    if (decodedKey === 'fields') {
      decodedValue = driveV2CompatTranslateFieldsExpression_(decodedValue, isTeamDriveList);
      return rawKey + '=' + encodeURIComponent(decodedValue);
    }

    return part;
  });

  return resourcePath + '?' + translatedParams.join('&');
}


/**
 * Drive API v3 스타일 검색식의 필드명을 v2 검색식으로 변환합니다.
 * 문자열 리터럴 안의 고객사명·폴더명은 절대 변경하지 않습니다.
 */
function driveV2CompatTranslateQueryExpression_(expression, isTeamDriveList) {
  var source = String(expression || '');
  if (isTeamDriveList) return source;

  var out = '';
  var token = '';
  var inQuote = false;
  var escaped = false;

  function flushToken() {
    if (!token) return;
    var translated = token;
    if (token === 'name') translated = 'title';
    if (token === 'createdTime') translated = 'createdDate';
    if (token === 'modifiedTime') translated = 'modifiedDate';
    out += translated;
    token = '';
  }

  for (var i = 0; i < source.length; i++) {
    var ch = source.charAt(i);

    if (inQuote) {
      flushToken();
      out += ch;
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === "'") {
        inQuote = false;
      }
      continue;
    }

    if (ch === "'") {
      flushToken();
      inQuote = true;
      out += ch;
      continue;
    }

    if (/[A-Za-z0-9_]/.test(ch)) {
      token += ch;
      continue;
    }

    flushToken();
    out += ch;
  }

  flushToken();
  return out;
}


function driveV2CompatTranslateFieldsExpression_(fieldsExpression, isTeamDriveList) {
  var fields = String(fieldsExpression || '');

  if (isTeamDriveList) {
    return fields
      .replace(/\bdrives\b/g, 'items')
      .replace(/\bfiles\b/g, 'items');
  }

  return fields
    .replace(/\bfiles\b/g, 'items')
    .replace(/\bname\b/g, 'title')
    .replace(/\bwebViewLink\b/g, 'alternateLink')
    .replace(/\bcreatedTime\b/g, 'createdDate')
    .replace(/\bmodifiedTime\b/g, 'modifiedDate');
}


function driveV2CompatSafeDecode_(value) {
  var source = String(value || '');
  try {
    return decodeURIComponent(source.replace(/\+/g, '%20'));
  } catch (err) {
    return source;
  }
}

function driveV2CompatPreparePayload_(payload) {
  if (!payload || typeof payload !== 'object') return payload;
  const copy = JSON.parse(JSON.stringify(payload));
  return driveV2CompatPayloadObject_(copy);
}

function driveV2CompatPayloadObject_(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(driveV2CompatPayloadObject_);
  if (Object.prototype.hasOwnProperty.call(obj, 'name') && !Object.prototype.hasOwnProperty.call(obj, 'title')) {
    obj.title = obj.name;
    delete obj.name;
  }
  if (Array.isArray(obj.parents)) {
    obj.parents = obj.parents.map(function(parent) {
      if (typeof parent === 'string') return { id: parent };
      return driveV2CompatPayloadObject_(parent);
    });
  }
  Object.keys(obj).forEach(function(key) {
    if (key === 'parents') return;
    obj[key] = driveV2CompatPayloadObject_(obj[key]);
  });
  return obj;
}

function driveV2CompatNormalizeResponse_(data) {
  if (!data || typeof data !== 'object') return data || {};
  const out = driveV2CompatNormalizeObject_(data);
  if (Array.isArray(out.items) && !out.files) {
    out.files = out.items.map(driveV2CompatNormalizeObject_);
  }
  if (Array.isArray(out.items) && !out.drives) {
    out.drives = out.items.map(function(item) {
      return item && typeof item === 'object'
        ? Object.assign({}, item, { id: item.id || '', name: item.name || item.title || '' })
        : item;
    });
  }
  return out;
}

function driveV2CompatNormalizeObject_(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(driveV2CompatNormalizeObject_);
  const out = {};
  Object.keys(obj).forEach(function(key) {
    out[key] = driveV2CompatNormalizeObject_(obj[key]);
  });
  if (out.title != null && out.name == null) out.name = out.title;
  if (out.alternateLink != null && out.webViewLink == null) out.webViewLink = out.alternateLink;
  if (out.createdDate != null && out.createdTime == null) out.createdTime = out.createdDate;
  if (out.modifiedDate != null && out.modifiedTime == null) out.modifiedTime = out.modifiedDate;
  return out;
}

function driveV2CompatFetch_(path, options) {
  const params = Object.assign(
    {
      method: 'get',
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    },
    options || {}
  );
  if (params.payload && typeof params.payload !== 'string') {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(driveV2CompatPreparePayload_(params.payload));
  } else if (params.payload && typeof params.payload === 'string' && /application\/json/i.test(String(params.contentType || ''))) {
    try { params.payload = JSON.stringify(driveV2CompatPreparePayload_(JSON.parse(params.payload))); } catch (e) {}
  }
  const res = UrlFetchApp.fetch(driveV2CompatBuildUrl_(path, false), params);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Drive API v2 오류 ' + code + ': ' + body);
  }
  return body ? driveV2CompatNormalizeResponse_(JSON.parse(body)) : {};
}
