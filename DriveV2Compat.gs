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
  let p = String(path || '');
  const isTeamDriveList = /^drives\b/.test(p);
  if (isTeamDriveList) {
    p = p.replace(/^drives\b/, 'teamdrives');
  }

  p = p
    .replace(/upload\/drive\/v3/g, 'upload/drive/v2')
    .replace(/drive\/v3/g, 'drive/v2')
    .replace(/fields=drives\(/g, 'fields=items(')
    .replace(/fields=files\(/g, 'fields=items(')
    .replace(/,files\(/g, ',items(')
    .replace(/files\(id,name/g, 'items(id,title')
    .replace(/\bname\b/g, 'title')
    .replace(/webViewLink/g, 'alternateLink')
    .replace(/createdTime/g, 'createdDate')
    .replace(/modifiedTime/g, 'modifiedDate');

  // Team Drive v2 resources still use "name", not "title".
  if (isTeamDriveList) {
    p = p.replace(/items\(id,title\)/g, 'items(id,name)');
    p = p.replace(/items\(id,title/g, 'items(id,name');
    p = p.replace(/title%20%3D/g, 'name%20%3D');
    p = p.replace(/title%3D/g, 'name%3D');
    p = p.replace(/title\+/g, 'name+');
  }
  return p;
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
