/**
 * SECURE ENTRY - Google Apps Script (Web App)
 * Actions:
 * - upload: store image into Drive folder, return photoId + photoUrl
 * - delete: trash file by photoId
 *
 * SETUP:
 * 1) Create folder in Drive (e.g. "Sensory Guard House Photos 2025")
 * 2) In Apps Script: Project Settings -> Script Properties:
 *    - PHOTO_FOLDER_ID = <Drive Folder ID>
 * 3) Deploy as Web App:
 *    - Execute as: Me
 *    - Who has access: Anyone (or Anyone with link)
 *
 * Note: Worker calls this server-to-server, so CORS isn't needed for browsers.
 */

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var action = String(body.action || "").toLowerCase();

    if (action === "upload") return handleUpload_(body);
    if (action === "delete") return handleDelete_(body);

    return json_(false, "Unknown action");
  } catch (err) {
    return json_(false, "Server error: " + (err && err.message ? err.message : err));
  }
}

function handleUpload_(body) {
  var folderId = PropertiesService.getScriptProperties().getProperty("PHOTO_FOLDER_ID");
  if (!folderId) return json_(false, "Missing script property: PHOTO_FOLDER_ID");

  var base64 = body.base64;
  var mimeType = body.mimeType || "image/jpeg";
  if (!base64) return json_(false, "Missing base64");

  var meta = body.meta || {};
  var docNo = (meta.docNo || "").toString().trim();
  var regNo = (meta.regNo || "").toString().trim();
  var name = (meta.name || "").toString().trim();

  // filename
  var ts = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyyMMdd_HHmmss");
  var safe = function(s){ return String(s||"").replace(/[^A-Za-z0-9_\-]/g,"").substring(0,24); };
  var fileName = "SE_" + ts + (docNo ? ("_DOC_" + safe(docNo)) : "") + (regNo ? ("_REG_" + safe(regNo)) : "") + (name ? ("_" + safe(name)) : "") + ".jpg";

  // decode
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(blob);

  // Sharing (anyone with link view) - optional; adjust if you prefer restricted
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // If domain policy blocks, still return fileId. Frontend can use alternate view route.
  }

  var photoId = file.getId();
  var photoUrl = file.getUrl(); // drive file url

  return json_(true, null, { photoId: photoId, photoUrl: photoUrl });
}

function handleDelete_(body) {
  var photoId = (body.photoId || "").toString().trim();
  if (!photoId) return json_(false, "Missing photoId");

  try {
    var file = DriveApp.getFileById(photoId);
    file.setTrashed(true);
    return json_(true, null, { photoId: photoId, trashed: true });
  } catch (err) {
    return json_(false, "Delete failed: " + (err && err.message ? err.message : err));
  }
}

function json_(ok, error, data) {
  var payload = { ok: !!ok };
  if (error) payload.error = error;
  if (data) {
    for (var k in data) payload[k] = data[k];
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
