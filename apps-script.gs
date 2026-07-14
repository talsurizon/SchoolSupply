/**
 * Google Apps Script for the SchoolSupply order sheet.
 * Paste into: spreadsheet > Extensions > Apps Script, replacing doPost.
 * IMPORTANT: after saving, go to Deploy > Manage deployments > ✎ Edit >
 * Version: "New version" > Deploy — otherwise the web app keeps running
 * the old code.
 *
 * To test from the editor, run testOrder() (▶), NOT doPost — doPost only
 * receives data via a real HTTP POST; running it directly has no event
 * object and throws on e.postData.
 *
 * Expected column order: timestamp | name | phone | email | order description | total
 * (adjust the appendRow below if your sheet's columns differ)
 */

// Hebrew labels for the per-item choices sent by the site (cart[].custom).
// Keys not listed here still appear, using the raw key as the label.
var CUSTOM_LABELS = {
  stickers: "מדבקות שם",
  ruler: "סרגל",
  scissors: "מספריים",
  headset: "אוזניות ועכבר אלחוטי",
};

function describeCart(cart) {
  return (cart || []).map(function (item) {
    var custom = item.custom || {};
    var choices = Object.keys(custom).map(function (key) {
      return (CUSTOM_LABELS[key] || key) + ": " + custom[key];
    }).join(" · ");
    var line = "כיתה " + item.letter + " (" + item.price + " ₪)";
    return choices ? line + " — " + choices : line;
  }).join("\n");
}

function doPost(e) {
  if (!e || !e.postData) {
    throw new Error("No POST data — to test from the editor, run testOrder() instead of doPost.");
  }
  var data = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.appendRow([
    new Date(data.timestamp), // real datetime value, shown in the sheet's timezone
    data.name,
    "'" + data.phone, // leading apostrophe forces text, keeping the leading 0
    data.email,
    describeCart(data.cart),
    data.total,
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Editor self-test: appends a clearly-marked row you can delete afterwards.
function testOrder() {
  doPost({
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        name: "TEST בדיקה - למחוק",
        phone: "0500000000",
        email: "test@test.com",
        total: 130,
        cart: [{
          letter: "ג'",
          price: 130,
          custom: { stickers: "קייפופ", ruler: "דונטס", scissors: "ורוד אפור", headset: "בלי" },
        }],
      }),
    },
  });
}
