// Google Apps Script backend for the beat store
// 1) Create a new Apps Script project at https://script.google.com
// 2) Paste this code into the editor
// 3) Deploy as a Web App: Execute as me, Who has access: Anyone
// 4) Copy the Web App URL into the site config as google_apps_script_url
// 5) Replace PASTE_YOUR_SPREADSHEET_ID_HERE with your Google Sheet ID

const DEFAULT_ADMIN_EMAIL = 'debeatjay@gmail.com';

function createCorsJsonOutput(payload) {
  const output = ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
  if (typeof output.setHeader === 'function') {
    output.setHeader('Access-Control-Allow-Origin', '*');
    output.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return output;
}

function doGet(e) {
  const params = e.parameter || {};
  const type = (params.type || '').toString().toLowerCase();

  if (type === 'offer_action' && params.token && params.action) {
    return handleOfferAction(params);
  }

  if (type === 'offer_lookup' && params.token) {
    return getOfferLookup(params);
  }

  return HtmlService.createHtmlOutput('Beat Store Apps Script is running.');
}

function normalizeFrontendUrl(value) {
  if (!value) {
    return '';
  }
  const raw = value.toString().trim();
  if (!raw) {
    return '';
  }
  return raw.replace(/\/+$/, '');
}

function getOfferLookup(params) {
  const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const offersSheet = getOrCreateSheet(ss, 'Offers');
  ensureHeaders(offersSheet, ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'itemId', 'customerEmail', 'adminEmail', 'scriptUrl', 'frontendUrl', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl']);

  const token = (params.token || params.offer_token || '').toString().trim();
  const rowIndex = findRowIndexByColumnValue(offersSheet, 'payLinkToken', token);
  if (rowIndex < 2) {
    return createCorsJsonOutput({ ok: false, message: 'Offer link not found' });
  }

  const values = offersSheet.getRange(rowIndex, 1, 1, offersSheet.getLastColumn()).getValues()[0];
  const headers = offersSheet.getDataRange().getValues()[0];
  const status = (values[headers.indexOf('status')] || '').toString().toLowerCase();

  if (status !== 'accepted') {
    return createCorsJsonOutput({ ok: false, message: 'Offer is not accepted yet', status: status || 'unknown' });
  }

  return createCorsJsonOutput({
    ok: true,
    offer: {
      token: token,
      beatTitle: values[headers.indexOf('beatTitle')] || '',
      itemId: values[headers.indexOf('itemId')] || '',
      offerPrice: values[headers.indexOf('offerPrice')] || '',
      customerEmail: values[headers.indexOf('customerEmail')] || values[headers.indexOf('email')] || '',
      adminEmail: values[headers.indexOf('adminEmail')] || '',
      frontendUrl: values[headers.indexOf('frontendUrl')] || '',
      payLinkUrl: values[headers.indexOf('payLinkUrl')] || '',
      checkoutUrl: values[headers.indexOf('payLinkUrl')] || '',
      cartResumeUrl: values[headers.indexOf('payLinkUrl')] || ''
    }
  });
}

function getWebAppUrl(params) {
  const scriptUrl = ((params && (params.scriptUrl || params.google_apps_script_url || params.googleAppsScriptUrl)) || '').toString().trim();
  if (scriptUrl) {
    return scriptUrl;
  }

  try {
    const service = ScriptApp.getService();
    if (service && typeof service.getUrl === 'function') {
      const currentUrl = service.getUrl();
      if (currentUrl) {
        return currentUrl.toString().trim();
      }
    }
  } catch (e) {
    // Fall through if service URL is unavailable
  }

  return '';
}

function generateSecureToken(length) {
  length = parseInt(length, 10) || 40;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  const seed = Utilities.getUuid() + Date.now().toString() + Math.random().toString();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  let token = '';
  for (let i = 0; i < length; i++) {
    token += alphabet[(digest[i % digest.length] + 256) % alphabet.length];
  }
  return token;
}

function findRowIndexByColumnValue(sheet, columnName, value) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return -1;
  }
  const headers = values[0] || [];
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex < 0) {
    return -1;
  }
  for (let row = 1; row < values.length; row++) {
    if ((values[row][columnIndex] || '').toString() === value.toString()) {
      return row + 1;
    }
  }
  return -1;
}

function renderOfferHtml(title, message, extraHtml) {
  return HtmlService.createHtmlOutput(
    `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title><style>body{font-family:Arial,Helvetica,sans-serif;background:#f4f7fb;color:#111;margin:0;padding:40px;} .container{max-width:640px;margin:0 auto;background:#fff;padding:32px;border-radius:16px;box-shadow:0 18px 50px rgba(15,23,42,0.12);} h1{font-size:24px;margin-bottom:16px;} p{line-height:1.7;color:#333;} a{color:#2563eb;text-decoration:none;}</style></head><body><div class="container"><h1>${title}</h1><p>${message}</p>${extraHtml || ''}</div></body></html>`
  );
}

function verifyFlutterwavePayment(transactionId, expectedAmount, expectedCurrency) {
  const secretKey = PropertiesService.getScriptProperties().getProperty('FLUTTERWAVE_SECRET_KEY');
  if (!secretKey) {
    throw new Error('Flutterwave secret key is not configured in Apps Script properties.');
  }
  if (!transactionId) {
    throw new Error('Flutterwave transaction ID is missing.');
  }

  const response = UrlFetchApp.fetch('https://api.flutterwave.com/v3/transactions/' + encodeURIComponent(transactionId) + '/verify', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + secretKey },
    muteHttpExceptions: true
  });
  const body = JSON.parse(response.getContentText() || '{}');
  const transaction = body.data || {};
  const amount = Number(expectedAmount || 0);
  const verified = body.status === 'success'
    && transaction.status === 'successful'
    && transaction.currency === expectedCurrency
    && Number(transaction.amount || 0) >= amount;

  if (!verified) {
    throw new Error('Flutterwave payment verification failed.');
  }
  return transaction;
}

function handleOfferAction(params) {
  const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const offersSheet = getOrCreateSheet(ss, 'Offers');
  ensureHeaders(offersSheet, ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'itemId', 'customerEmail', 'adminEmail', 'scriptUrl', 'frontendUrl', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl']);

  const token = (params.token || '').toString().trim();
  const action = (params.action || '').toString().toLowerCase();
  const rowIndex = findRowIndexByColumnValue(offersSheet, 'actionToken', token);
  if (rowIndex < 2) {
    return renderOfferHtml('Offer action not found', 'This action link is invalid or has expired.');
  }

  const values = offersSheet.getRange(rowIndex, 1, 1, offersSheet.getLastColumn()).getValues()[0];
  const headers = offersSheet.getDataRange().getValues()[0];
  const statusIndex = headers.indexOf('status');
  const currentStatus = (values[statusIndex] || '').toString().toLowerCase() || 'pending';
  if (currentStatus !== 'pending') {
    return renderOfferHtml('Offer already processed', `This offer has already been ${currentStatus}.`);
  }

  const customerEmail = (values[headers.indexOf('customerEmail')] || values[headers.indexOf('email')] || '').toString().trim();
  const adminEmail = (values[headers.indexOf('adminEmail')] || '').toString().trim();
  const beatTitle = (values[headers.indexOf('beatTitle')] || '').toString();
  const beatGenre = (values[headers.indexOf('beatGenre')] || '').toString();
  const beatBpm = (values[headers.indexOf('beatBpm')] || '').toString();
  const beatKey = (values[headers.indexOf('beatKey')] || '').toString();
  const offerPrice = (values[headers.indexOf('offerPrice')] || '').toString();
  const offerMessage = (values[headers.indexOf('offerMessage')] || '').toString();
  const itemId = (values[headers.indexOf('itemId')] || '').toString();
  const scriptUrl = getWebAppUrl({ scriptUrl: values[headers.indexOf('scriptUrl')] || '' });
  const fallbackFrontendUrl = normalizeFrontendUrl(PropertiesService.getScriptProperties().getProperty('STORE_FRONTEND_URL'));
  const rowFrontendUrl = normalizeFrontendUrl(values[headers.indexOf('frontendUrl')] || '');
  const frontendUrl = rowFrontendUrl || fallbackFrontendUrl;
  const actionTimestamp = new Date().toISOString();

  if (action === 'accept') {
    const payLinkToken = generateSecureToken(48);
    const payLinkUrl = frontendUrl
      ? `${frontendUrl}${frontendUrl.indexOf('?') >= 0 ? '&' : '?'}offer_token=${encodeURIComponent(payLinkToken)}`
      : `${scriptUrl}?type=offer_pay&token=${encodeURIComponent(payLinkToken)}`;

    if (!frontendUrl && !scriptUrl) {
      return renderOfferHtml('Offer accepted blocked', 'This seller accepted offer cannot continue because the website frontend URL is not configured.');
    }

    if (frontendUrl) {
      offersSheet.getRange(rowIndex, headers.indexOf('frontendUrl') + 1).setValue(frontendUrl);
    }

    offersSheet.getRange(rowIndex, statusIndex + 1).setValue('accepted');
    offersSheet.getRange(rowIndex, headers.indexOf('actionTaken') + 1).setValue('accept');
    offersSheet.getRange(rowIndex, headers.indexOf('actionTimestamp') + 1).setValue(actionTimestamp);
    offersSheet.getRange(rowIndex, headers.indexOf('payLinkToken') + 1).setValue(payLinkToken);
    offersSheet.getRange(rowIndex, headers.indexOf('payLinkUrl') + 1).setValue(payLinkUrl);

    if (customerEmail) {
      const body = [`<p>Hi there,</p>`, `<p>Great news! Your offer for <strong>${escapeHtml(beatTitle)}</strong> has been accepted by the seller.</p>`, `<p><strong>Accepted offer amount:</strong> ₦${escapeHtml(offerPrice)}</p>`, `<p>Use the cart link below to continue your purchase on the website with the newly agreed amount.</p>`, `<p><a href="${payLinkUrl}" target="_blank" style="background:#2563eb;color:#fff;padding:12px 18px;border-radius:8px;display:inline-block;text-decoration:none;">Open cart and continue purchase</a></p>`, `<p>This is your unique offer checkout link. It loads the beat in the website cart and continues the normal license checkout flow.</p>`, `<p>If you did not make this offer, please reply to this email immediately.</p>`].join('');
      sendNotificationEmail({
        to: customerEmail,
        replyTo: adminEmail || 'debeatjay@gmail.com',
        subject: 'Your offer has been accepted — continue cart checkout',
        htmlBody: body
      });
    }

    if (adminEmail) {
      const body = [`<p>Hi,</p>`, `<p>The offer from <strong>${escapeHtml(customerEmail)}</strong> for <strong>${escapeHtml(beatTitle)}</strong> has been accepted.</p>`, `<p><strong>Offer price:</strong> ₦${escapeHtml(offerPrice)}</p>`, `<p><strong>Item ID:</strong> ${escapeHtml(itemId)}</p>`, `<p>The paylink has been generated and sent to the customer.</p>`].join('');
      sendNotificationEmail({
        to: adminEmail,
        replyTo: customerEmail || '',
        subject: `Offer accepted for ${beatTitle}`,
        htmlBody: body
      });
    }

    return renderOfferHtml('Offer accepted', `The offer has been accepted and a cart link has been sent to ${escapeHtml(customerEmail)}.`, `<p><strong>Cart URL:</strong> <a href="${payLinkUrl}" target="_blank">${payLinkUrl}</a></p>`);
  }

  if (action === 'decline') {
    offersSheet.getRange(rowIndex, statusIndex + 1).setValue('declined');
    offersSheet.getRange(rowIndex, headers.indexOf('actionTaken') + 1).setValue('decline');
    offersSheet.getRange(rowIndex, headers.indexOf('actionTimestamp') + 1).setValue(actionTimestamp);

    if (customerEmail) {
      const body = [`<p>Hi there,</p>`, `<p>We wanted to let you know that your offer for <strong>${escapeHtml(beatTitle)}</strong> was declined.</p>`, `<p>Your offer amount was: ₦${escapeHtml(offerPrice)}</p>`, `<p>Please feel free to submit a new offer or select another beat.</p>`].join('');
      sendNotificationEmail({
        to: customerEmail,
        replyTo: adminEmail || 'debeatjay@gmail.com',
        subject: 'Your offer was declined',
        htmlBody: body
      });
    }

    return renderOfferHtml('Offer declined', `The offer has been declined and the customer has been notified.`);
  }

  return renderOfferHtml('Invalid action', 'The requested action is not supported.');
}

function redirectOfferCartFromToken(token) {
  const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const offersSheet = getOrCreateSheet(ss, 'Offers');
  ensureHeaders(offersSheet, ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'itemId', 'customerEmail', 'adminEmail', 'scriptUrl', 'frontendUrl', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl']);

  const rowIndex = findRowIndexByColumnValue(offersSheet, 'payLinkToken', token);
  if (rowIndex < 2) {
    return renderOfferHtml('Offer checkout link not found', 'This offer cart link is invalid or has expired.');
  }

  const values = offersSheet.getRange(rowIndex, 1, 1, offersSheet.getLastColumn()).getValues()[0];
  const headers = offersSheet.getDataRange().getValues()[0];
  const status = (values[headers.indexOf('status')] || '').toString().toLowerCase();
  if (status !== 'accepted') {
    return renderOfferHtml('Offer checkout unavailable', 'This offer is not currently accepted.');
  }

  const payLinkUrl = (values[headers.indexOf('payLinkUrl')] || '').toString().trim();
  if (!payLinkUrl) {
    return renderOfferHtml('Offer cart not found', 'The accepted offer cart link is missing.');
  }

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Offer checkout</title></head><body><p>Loading your cart...</p><script>window.location.replace(${JSON.stringify(payLinkUrl)});</script></body></html>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderPayLinkPage(token) {
  const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const offersSheet = getOrCreateSheet(ss, 'Offers');
  ensureHeaders(offersSheet, ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'itemId', 'customerEmail', 'adminEmail', 'scriptUrl', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl']);

  const rowIndex = findRowIndexByColumnValue(offersSheet, 'payLinkToken', token);
  if (rowIndex < 2) {
    return renderOfferHtml('Paylink not found', 'This payment link is invalid or has already been used.');
  }

  const values = offersSheet.getRange(rowIndex, 1, 1, offersSheet.getLastColumn()).getValues()[0];
  const headers = offersSheet.getDataRange().getValues()[0];
  const status = (values[headers.indexOf('status')] || '').toString().toLowerCase();
  if (status !== 'accepted') {
    return renderOfferHtml('Payment link unavailable', 'This payment link is not active because the offer is not currently accepted.');
  }

  const beatTitle = (values[headers.indexOf('beatTitle')] || '').toString();
  const offerPrice = (values[headers.indexOf('offerPrice')] || '').toString();
  const itemId = (values[headers.indexOf('itemId')] || '').toString();
  const customerEmail = (values[headers.indexOf('customerEmail')] || values[headers.indexOf('email')] || '').toString();
  const adminEmail = (values[headers.indexOf('adminEmail')] || DEFAULT_ADMIN_EMAIL).toString().trim();
  const amountInKobo = Math.round((parseFloat(offerPrice) || 0) * 100);

  const pageHtml = renderOfferHtml('Checkout unavailable', 'Online checkout is no longer available. Please contact the seller to complete this order.');

  return HtmlService.createHtmlOutput(pageHtml).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderPayLinkPage_v2(token) {
  const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const offersSheet = getOrCreateSheet(ss, 'Offers');
  ensureHeaders(offersSheet, ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'itemId', 'customerEmail', 'adminEmail', 'scriptUrl', 'frontendUrl', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl']);

  const rowIndex = findRowIndexByColumnValue(offersSheet, 'payLinkToken', token);
  if (rowIndex < 2) {
    return renderOfferHtml('Paylink not found', 'This payment link is invalid or has already been used.');
  }

  const values = offersSheet.getRange(rowIndex, 1, 1, offersSheet.getLastColumn()).getValues()[0];
  const headers = offersSheet.getDataRange().getValues()[0];
  const status = (values[headers.indexOf('status')] || '').toString().toLowerCase();
  if (status !== 'accepted') {
    return renderOfferHtml('Payment link unavailable', 'This payment link is not active because the offer is not currently accepted.');
  }

  const beatTitle = (values[headers.indexOf('beatTitle')] || '').toString();
  const offerPrice = (values[headers.indexOf('offerPrice')] || '').toString();
  const itemId = (values[headers.indexOf('itemId')] || '').toString();
  const customerEmail = (values[headers.indexOf('customerEmail')] || values[headers.indexOf('email')] || '').toString();
  const adminEmail = (values[headers.indexOf('adminEmail')] || DEFAULT_ADMIN_EMAIL).toString().trim();
  const amountInKobo = Math.round((parseFloat(offerPrice) || 0) * 100);

  // Build HTML in lines to avoid long single-line template issues
  const pageLines = [];
  pageLines.push('<!DOCTYPE html>');
  pageLines.push('<html>');
  pageLines.push('<head>');
  pageLines.push('<meta charset="utf-8">');
  pageLines.push('<meta name="viewport" content="width=device-width,initial-scale=1">');
  pageLines.push('<title>Secure payment link</title>');
  pageLines.push('<style>body{font-family:Arial,Helvetica,sans-serif;background:#f4f7fb;color:#111;margin:0;padding:24px;} .container{max-width:720px;margin:0 auto;background:#fff;padding:28px;border-radius:18px;box-shadow:0 18px 50px rgba(15,23,42,0.12);} h1{font-size:24px;margin-bottom:16px;} p{line-height:1.75;color:#333;} .details{margin:20px 0;padding:18px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;} .button{background:#2563eb;color:#fff;padding:14px 20px;border:none;border-radius:12px;font-size:16px;cursor:pointer;text-decoration:none;} .note{margin-top:18px;color:#475569;font-size:14px;}</style>');
  pageLines.push('</head>');
  pageLines.push('<body>');
  pageLines.push('<div class="container">');
  pageLines.push('<h1>Secure payment</h1>');
  pageLines.push('<p>Hi ' + escapeHtml(customerEmail || 'Customer') + ',</p>');
  pageLines.push('<p>Your secure payment link for <strong>' + escapeHtml(beatTitle) + '</strong> is ready.</p>');
  pageLines.push('<div class="details">');
  pageLines.push('<p><strong>Item ID:</strong> ' + escapeHtml(itemId) + '</p>');
  pageLines.push('<p><strong>Accepted offer price:</strong> ₦' + escapeHtml(offerPrice) + '</p>');
  pageLines.push('<p><strong>Email:</strong> ' + escapeHtml(customerEmail) + '</p>');
  pageLines.push('</div>');
  pageLines.push('<button id="pay-button" class="button">Open secure payment page</button>');
  pageLines.push('<p id="status-message" class="note">Online checkout is no longer available. Please contact the seller to complete this order.</p>');
  pageLines.push('</div>');
  pageLines.push('');

  // Legacy offer-page client code retained for compatibility with old links.
  pageLines.push('<script>');
  pageLines.push('(function(){');
  pageLines.push('  const payButton = document.getElementById("pay-button");');
  pageLines.push('  const statusMessage = document.getElementById("status-message");');
  pageLines.push('  const apiUrl = window.location.origin + window.location.pathname;');
  pageLines.push('  function serializeFormParams(params) {');
  pageLines.push('    return Object.keys(params).map(function(key) { return encodeURIComponent(key) + "=" + encodeURIComponent(params[key] || ""); }).join("&");');
  pageLines.push('  }');
  pageLines.push('  function submitFallbackForm(url, params) {');
  pageLines.push('    const form = document.createElement("form");');
  pageLines.push('    form.method = "POST";');
  pageLines.push('    form.action = url;');
  pageLines.push('    form.target = "_blank";');
  pageLines.push('    form.style.display = "none";');
  pageLines.push('    Object.keys(params).forEach(function(key) {');
  pageLines.push('      const input = document.createElement("input");');
  pageLines.push('      input.type = "hidden";');
  pageLines.push('      input.name = key;');
  pageLines.push('      input.value = params[key] || "";');
  pageLines.push('      form.appendChild(input);');
  pageLines.push('    });');
  pageLines.push('    document.body.appendChild(form);');
  pageLines.push('    form.submit();');
  pageLines.push('  }');
  pageLines.push('  function sendPaymentToServer(payload) {');
  pageLines.push('    const body = serializeFormParams(payload);');
  pageLines.push('    return fetch(apiUrl, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body, cache: "no-store" })');
  pageLines.push('      .then(function(res) { return res.text().then(function(text) { if (!res.ok) { throw new Error("HTTP " + res.status + " " + res.statusText + " — " + text); } try { return JSON.parse(text); } catch (parseError) { throw new Error("Invalid JSON response: " + text); } }); })');
  pageLines.push('      .catch(function(err) {');
  pageLines.push('        console.warn("Payment sync failed, using form fallback", err);');
  pageLines.push('        submitFallbackForm(apiUrl, payload);');
  pageLines.push('        throw err;');
  pageLines.push('      });');
  pageLines.push('  }');
  pageLines.push('  payButton.addEventListener("click", function(){');
  pageLines.push('    alert("Online checkout is unavailable. Please contact the seller to complete this order."); return;');
  pageLines.push('    payButton.disabled = true; payButton.textContent = "Opening payment...";');
  pageLines.push('    return; const handler = PaymentWidget.setup({');
  pageLines.push('      key: "",');
  pageLines.push('      email: ' + JSON.stringify(customerEmail) + ',');
  pageLines.push('      amount: ' + amountInKobo + ',');
  pageLines.push('      currency: "NGN",');
  pageLines.push('      metadata: { custom_fields: [' + '{display_name:"Offer Token",variable_name:"offer_token",value:' + JSON.stringify(token) + '},' + '{display_name:"Item ID",variable_name:"item_id",value:' + JSON.stringify(itemId) + '}' + '] },');
  pageLines.push('      callback: function(response){');
  pageLines.push('        statusMessage.textContent = "Payment completed. Sending confirmation email...";');
  pageLines.push('        try {');
  pageLines.push('          const payload = {');
  pageLines.push('            type: "payment",');
  pageLines.push('            name: ' + JSON.stringify('Customer') + ',');
  pageLines.push('            email: ' + JSON.stringify(customerEmail) + ',');
  pageLines.push('            customerEmail: ' + JSON.stringify(customerEmail) + ',');
  pageLines.push('            adminEmail: ' + JSON.stringify(adminEmail) + ',');
  pageLines.push('            paymentReference: response.reference || "",');
  pageLines.push('            amount: ' + JSON.stringify(offerPrice) + ',');
  pageLines.push('            currency: "NGN",');
  pageLines.push('            status: "success",');
  pageLines.push('            orderItems: JSON.stringify([{ beat: ' + JSON.stringify(beatTitle) + ', license: "Accepted offer", price: ' + JSON.stringify(offerPrice) + ', itemId: ' + JSON.stringify(itemId) + ' }]),');
  pageLines.push('            orderSummary: JSON.stringify({ items: [{ beat: ' + JSON.stringify(beatTitle) + ', price: ' + JSON.stringify(offerPrice) + ' }], total: ' + JSON.stringify(offerPrice) + ' }),');
  pageLines.push('            downloadLinks: JSON.stringify([]),');
  pageLines.push('            rawResponse: JSON.stringify(response)');
  pageLines.push('          };');
  pageLines.push('          sendPaymentToServer(payload)');
  pageLines.push('            .then(function(body){');
  pageLines.push('              statusMessage.textContent = "Payment recorded. Confirmation email sent.";');
  pageLines.push('              payButton.textContent = "Payment complete";');
  pageLines.push('            })');
  pageLines.push('            .catch(function(err){');
  pageLines.push('              statusMessage.textContent = "Payment completed. Server sync failed; opening fallback tab to submit payment. If your payment is not recorded, follow the fallback link.";');
  pageLines.push('              payButton.textContent = "Retry sync?";');
  pageLines.push('            });');
  pageLines.push('        } catch(err) { statusMessage.textContent = "Error preparing server notification: " + err.message; }');
  pageLines.push('      },');
  pageLines.push('      onClose: function(){ statusMessage.textContent = "Payment window closed. If you completed payment, refresh this page."; payButton.disabled = false; payButton.textContent = "Open secure payment page"; }');
  pageLines.push('    });');
  pageLines.push('    handler.openIframe();');
  pageLines.push('  });');
  pageLines.push('})();');
  pageLines.push('</script>');
  pageLines.push('</body>');
  pageLines.push('</html>');

  return HtmlService.createHtmlOutput(pageLines.join('\n')).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const normalizedExisting = existingHeaders.map(function (value) {
    return (value || '').toString().trim();
  });

  for (let i = 0; i < headers.length; i += 1) {
    const desiredHeader = headers[i];
    if (normalizedExisting.indexOf(desiredHeader) >= 0) {
      continue;
    }

    const priorDesiredHeaders = headers.slice(0, i).filter(function (candidate) {
      return normalizedExisting.indexOf(candidate) >= 0;
    });

    let insertAfterColumn = 0;
    if (priorDesiredHeaders.length > 0) {
      const latestPrior = priorDesiredHeaders[priorDesiredHeaders.length - 1];
      const priorIndex = normalizedExisting.indexOf(latestPrior);
      if (priorIndex >= 0) {
        insertAfterColumn = priorIndex + 1;
      }
    }

    sheet.insertColumnAfter(insertAfterColumn);
    const targetColumn = insertAfterColumn + 1;
    sheet.getRange(1, targetColumn).setValue(desiredHeader);

    normalizedExisting.splice(targetColumn - 1, 0, desiredHeader);
  }
}

function hashPassword(password) {
  const value = password || '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  return Utilities.base64Encode(digest);
}

function getAccountSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'Accounts');
  ensureHeaders(sheet, ['timestamp', 'type', 'name', 'email', 'password', 'status']);
  return sheet;
}

function findAccountByEmail(sheet, email) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return null;
  }

  const normalizedEmail = (email || '').toString().trim().toLowerCase();
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const rowEmail = (row[3] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      return {
        rowIndex: i + 1,
        name: row[2] || '',
        email: row[3] || '',
        password: row[4] || '',
        status: row[5] || ''
      };
    }
  }

  return null;
}

function updateAccountPassword(sheet, email, newPassword) {
  const account = findAccountByEmail(sheet, email);
  if (!account || !account.rowIndex) {
    return false;
  }

  sheet.getRange(account.rowIndex, 5).setValue(hashPassword(newPassword));
  return true;
}

function sendNotificationEmail(options) {
  if (!options || !options.to) {
    return { ok: false, reason: 'missing-recipient' };
  }

  const mailOptions = {
    to: options.to,
    replyTo: options.replyTo || 'debeatjay@gmail.com',
    subject: options.subject,
    htmlBody: options.htmlBody
  };

  if (Array.isArray(options.attachments) && options.attachments.length) {
    mailOptions.attachments = options.attachments;
  }

  MailApp.sendEmail(mailOptions);

  return { ok: true, to: options.to };
}

function tryParseJson(value) {
  try {
    return JSON.parse(value);
  } catch (e) {
    return null;
  }
}

function findLicenseTemplate(licenseKey) {
  const key = (licenseKey || '').toString().trim().toLowerCase();
  if (key.indexOf('mp3') >= 0) {
    return {
      title: 'MP3 License',
      content: [
        'This Non-Exclusive MP3 LICENSE (Contract Preview Only) License Agreement (the "Agreement"), having been made on and effective as of [Effective Date] by and between Preview Only p/k/a Preview Only (the “Producer” or “Licensor”) and Licensee, sets forth the terms and conditions of Licensee’s use, and the rights granted in, the Producer’s instrumental music file entitled [Beat Title] (the "Beat") in consideration for Licensee’s payment of ₦38,000 (the "License Fee") on a so-called "MP3 LICENSE (Contract Preview Only)" basis.',
        '',
        'This Agreement is issued solely in connection with and for Licensee’s use of the Beat pursuant and subject to all terms and conditions set forth herein.',
        '',
        '1. License Fee: The Licensee shall make payment of the License Fee to Licensor on the date of this Agreement. All rights granted to Licensee by Producer in the Beat are conditional upon Licensee’s timely payment of the License Fee. The License Fee is a one-time payment for the rights granted to Licensee and this Agreement is not valid until the License Fee has been paid.',
        '',
        '2. Delivery of the Beat:',
        '   - Licensor agrees to deliver the Beat as a high-quality MP3, as such terms are understood in the music industry.',
        '   - Licensor shall use commercially reasonable efforts to deliver the Beat to Licensee immediately after payment of the License Fee is made. Licensee will receive the Beat via email to the email address Licensee provided to Licensor.',
        '',
        '3. Term: The Term of this Agreement shall be ten (10) years and this license shall expire on the ten (10) year anniversary of the Effective Date.',
        '',
        '4. Use of the Beat:',
        '   - In consideration for Licensee’s payment of the License Fee, the Producer hereby grants Licensee a limited non-exclusive, nontransferable license and the right to incorporate, include and/or use the Beat in the preparation of one (1) new song or to incorporate the Beat into a new piece of instrumental music created by the Licensee. Licensee may create the new song or new instrumental music by recording his/her written lyrics over the Beat and/or by incorporating portions/samples of the Beat into pre-existing instrumental music written, produced and/or owned by Licensee. The new song or piece of instrumental music created by Licensee which incorporates some or all of the Beat shall be referred to as the "New Song". Permission is granted to Licensee to modify the arrangement, length, tempo, or pitch of the Beat in preparation of the New Song for public release.',
        '   - This License grants Licensee a worldwide, non-exclusive license to use the Beat as incorporated in the New Song in the manners and for the purposes expressly provided for herein, subject to the sale restrictions, limitations and prohibited uses stated in this Agreement. Licensee acknowledges and agrees that any and all rights granted to Licensee in the Beat pursuant to this Agreement are on a NON-EXCLUSIVE basis and Producer shall continue to license the Beat upon the same or similar terms and conditions as this Agreement to other potential third-party licensees.',
        '',
        'Permitted uses:',
        '   - The New Song may be used for any promotional purposes, including but not limited to, a release in a single format, for inclusion in a mixtape or free compilation of music bundled together (EP or album), and/or promotional, non-monetized digital streaming.',
        '   - Licensee may perform the song publicly for-profit performances and for unlimited non-profit performances, including but not limited to, at a live performance (i.e. concert, festival, nightclub etc.), on terrestrial or satellite radio, and/or on the internet via third-party streaming services (Spotify, YouTube, iTunes Radio etc.). The New Song may be played on 2 terrestrial or satellite radio stations.',
        '   - The Licensee may use the New Song in synchronization with One (1) audiovisual work no longer than five (5) minutes in length (a "Video"). In the event that the New Song itself is longer than five (5) minutes in length, the Video may not play for longer than the length of the New Song. The Video may be broadcast on any television network and/or uploaded to the internet for digital streaming and/or free download by the public including but not limited to on YouTube and/or Vevo. Producer grants no other synchronization rights to Licensee.',
        '   - The Licensee may make the New Song available for sale in physical and/or digital form and sell 0 downloads/physical music products and are allowed 50,000 monetized audio streams, 0 monetized video streams, and are allowed unlimited free downloads. The New Song may be available for sale as a single and/or included in a compilation of other songs bundled together by Licensee as an EP or a full-length Album. The New Song may be sold via digital retailers for permanent digital download in mp3 format and/or physical format, including compact disc and vinyl records. For clarity and avoidance of doubt, the Licensee does NOT have the right to sell the Beat in the form that it was delivered to Licensee. The Licensee must create a New Song (or instrumental as detailed above) for its rights under this provision to vest. Any sale of the Beat in its original form by Licensee shall be a material breach of this Agreement and the Licensee shall be liable to the Licensor for damages as provided hereunder.',
        '',
        '5. Subject to the Licensee’s compliance with the terms and conditions of this Agreement, Licensee shall not be required to account or pay to Producer any royalties, fees, or monies paid to or collected by the Licensee (expressly excluding mechanical royalties), or which would otherwise be payable to Producer in connection with the use/exploitation of the New Song as set forth in this Agreement.',
        '',
        '6. Restrictions on the Use of the Beat:',
        '   - The rights granted to Licensee are NON-TRANSFERABLE and Licensee may not transfer or assign any of its rights hereunder to any third-party.',
        '   - The Licensee shall not synchronize, or permit third parties to synchronize, the Beat or New Song with any audiovisual works EXCEPT as expressly provided for and pursuant to Paragraph 4(b)(iii) of this Agreement for use in one (1) Video. This restriction includes, but is not limited to, use of the Beat and/or New Song in television, commercials, film/movies, theatrical works, video games, and in any other form on the Internet which is not expressly permitted herein.',
        '   - The Licensee shall not have the right to license or sublicense any use of the Beat or of the New Song, in whole or in part, for any so-called "samples".',
        '   - Licensee shall not engage in any unlawful copying, streaming, duplicating, selling, lending, renting, hiring, broadcasting, uploading, or downloading to any database, servers, computers, peer to peer sharing, or other file-sharing services, posting on websites, or distribution of the Beat in the form, or a substantially similar form, as delivered to Licensee. Licensee may send the Beat file to any individual musician, engineer, studio manager or other people who are working on the New Song.',
        '   - THE LICENSEE IS EXPRESSLY PROHIBITED FROM REGISTERING THE BEAT AND/OR NEW SONG WITH ANY CONTENT IDENTIFICATION SYSTEM, SERVICE PROVIDER, MUSIC DISTRIBUTOR, RECORD LABEL OR DIGITAL AGGREGATOR (for example TuneCore or CDBaby, and any other provider of user-generated content identification services). The purpose of this restriction is to prevent you from receiving a copyright infringement takedown notice from a third party who also received a non-exclusive license to use the Beat in a New Song. The Beat has already been tagged for Content Identification by Producer as a pre-emptive measure to protect all interested parties in the New Song. If you do not adhere to this policy, you are in violation of the terms of this License and your license to use the Beat and/or New Song may be revoked without notice or compensation to you.',
        '   - As applicable to both the underlying composition in the Beat and to the master recording of the Beat: (i) The parties acknowledge and agree that the New Song is a "derivative work", as that term is used in the United States Copyright Act; (ii) As applicable to the Beat and/or the New Song, there is no intention by the parties to create a joint work; and (iii) There is no intention by the Licensor to grant any rights in and/or to any other derivative works that may have been created by other third-party licensees.',
        '',
        '7. Ownership:',
        '   - The Producer is and shall remain the sole owner and holder of all rights, title, and interest in the Beat, including all copyrights to and in the sound recording and the underlying musical compositions written and composed by Producer. Nothing contained herein shall constitute an assignment by Producer to Licensee of any of the foregoing rights. Licensee may not, under any circumstances, register or attempt to register the New Song and/or the Beat with the U.S. Copyright Office. The aforementioned right to register the New Song and/or the Beat shall be strictly limited to Producer. Licensee will, upon request, execute, acknowledge and deliver to Producer such additional documents as Producer may deem necessary to evidence and effectuate Producer’s rights hereunder, and Licensee hereby grants to Producer the right as attorney-in-fact to execute, acknowledge, deliver and record in the U.S. Copyright Office or elsewhere any and all such documents if Licensee shall fail to execute same within five (5) days after so requested by Producer.',
        '   - For the avoidance of doubt, you do not own the master or the sound recording rights in the New Song. You have been licensed the right to use the Beat in the New Song and to commercially exploit the New Song based on the terms and conditions of this Agreement.',
        '   - Notwithstanding the above, you do own the lyrics or other original musical components of the New Song that were written or composed solely by you.',
        '   - With respect to the publishing rights and ownership of the underlying composition embodied in the New Song, the Licensee, and the Producer hereby acknowledge and agree that the underlying composition shall be owned/split between them as follows:.',
        '   - Writer Share Licensor Name (Licensor) 50% Licensee Name (Licensee) 50%',
        '   - Producer shall own, control, and administer Fifty Percent (50%) of the so-called "Publishers Share" of the underlying composition.',
        '   - In the event that Licensee wishes to register his/her interests and rights to the underlying composition of the New Song with their Performing Rights Organization ("PRO"), Licensee must simultaneously identify and register the Producers share and ownership interest in the composition to indicate that Producer wrote and owns 50% of the composition in the New Song and as the owner of 50% of the Publishers share of the New Song.',
        '1. - The licensee shall be deemed to have signed, affirmed and ratified its acceptance of the terms of this Agreement by virtue of its payment of the License Fee to Licensor and its electronic acceptance of its terms and conditions at the time Licensee made payment of the License Fee.',
        '',
        '2. -Mechanical License: If any selection or musical composition, or any portion thereof, recorded in the New Song hereunder is written or composed by Producer, in whole or in part, alone or in collaboration with others, or is owned or controlled, in whole or in part, directly or indirectly, by Producer or any person, firm, or corporation in which Producer has a direct or indirect interest, then such selection and/or musical composition shall be hereinafter referred to as a "Controlled Composition". Producer hereby agrees to issue or cause to be issued, as applicable, to Licensee, mechanical licenses in respect of each Controlled Composition, which are embodied on the New Song. For that license, on the United States and Canada sales, Licensee will pay mechanical royalties at one hundred percent (100%) of the minimum statutory rate, subject to no cap of that rate for albums and/or EPs. For license outside the United States and Canada, the mechanical royalty rate will be the rate prevailing on an industry-wide basis in the country concerned on the date that this agreement has been entered into.',
        '',
        '3. -Credit: Licensee shall have the right to use and permit others to use Producers approved name, approved likeness, and other approved identification and approved biographical material concerning the Producer solely for purposes of trade and otherwise without restriction solely in connection with the New Song recorded hereunder. Licensee shall use best efforts to have Producer credited as a "producer" and shall give Producer appropriate production and songwriting credit on all compact discs, record, music video, and digital labels or any other record configuration manufactured which is now known or created in the future that embodies the New Song created hereunder and on all cover liner notes, any records containing the New Song and on the front and/or back cover of any album listing the New Song and other musician credits. The licensee shall use its best efforts to ensure that Producer is properly credited and Licensee shall check all proofs for the accuracy of credits, and shall use its best efforts to cure any mistakes regarding Producers credit. In the event of any failure by Licensee to issue the credit to Producer, Licensee must use reasonable efforts to correct any such failure immediately and on a prospective basis. Such credit shall be in the substantial form: "Produced by Preview Only.',
        '',
        '4. -Licensors Option: Licensor shall have the option, at Licensors sole discretion, to terminate this License at any time within three (3) years of the date of this Agreement upon written notice to Licensee. In the event that Licensor exercises this option, Licensor shall pay to Licensee a sum equal to Two Hundred Percent (200%) of the License Fee paid by Licensee. Upon Licensors exercise of the option, Licensee must immediately remove the New Song from any and all digital and physical distribution channels and must immediately cease access to any streams and/or downloads of the New Song by the general public.',
        '',
        '5. -Breach by Licensee: The licensee shall have five (5) business days from its receipt of written notice by Producer and/or Producers authorized representative to cure any alleged breach of this Agreement by Licensee. Licensees failure to cure the alleged breach within five (5) business days shall result in Licensees default of its obligations, its breach of this Agreement, and at Producers sole discretion, the termination of Licensees rights hereunder. If Licensee engages in the commercial exploitation and/or sale of the Beat or New Song outside of the manner and amount expressly provided for in this Agreement, Licensee shall be liable to Producer for monetary damages in an amount equal to any and all monies paid, collected by, or received by Licensee, or any third party on its behalf, in connection with such unauthorized commercial exploitation of the Beat and/or New Song. Licensee recognizes and agrees that a breach or threatened breach of this Agreement by Licensee give rise to irreparable injury to Producer, which may not be adequately compensated by damages. Accordingly, in the event of a breach or threatened breach by the Licensee of the provisions of this Agreement, Producer may seek and shall be entitled to a temporary restraining order and a preliminary injunction restraining the Licensee from violating the provisions of this Agreement. Nothing herein shall prohibit Producer from pursuing any other available legal or equitable remedy from such breach or threatened breach, including but not limited to the recovery of damages from Licensee. The Licensee shall be responsible for all costs, expenses or damages that Producer incurs as a result of any violation by the Licensee of any provision of this Agreement. Licensee obligation shall include court costs, litigation expenses, and reasonable attorneys fees.',
        '',
        'Warranties, Representations, and Indemnification: Licensee hereby agrees that Licensor has not made any guarantees or promises that the Beat fits the particular creative use or musical purpose intended or desired by the Licensee. The Beat, its sound recording, and the underlying musical composition embodied therein are licensed to the Licensee "as is" without warranties of any kind or fitness for a particular purpose.',
        '',
        '1. Parties hereto shall indemnify and hold each other harmless from any and all third party claims, liabilities, costs, losses, damages or expenses as are actually incurred by the non-defaulting party and shall hold the non-defaulting party, free, safe, and harmless against and from any and all claims, suits, demands, costs, liabilities, loss, damages, judgments, recoveries, costs, and expenses; (including, without limitation, reasonable attorneys fees), which may be made or brought, paid, or incurred by reason of any breach or claim of breach of the warranties and representations hereunder by the defaulting party, their agents, heirs, successors, assigns and employees, which have been reduced to final judgment; provided that prior to final judgment, arising out of any breach of any representations or warranties of the defaulting party contained in this agreement or any failure by defaulting party to perform any obligations on its part to be performed hereunder the non-defaulting party has given the defaulting party prompt written notice of all claims and the right to participate in the defense with counsel of its choice at its sole expense. In no event shall artist be entitled to seek injunctive or any other equitable relief for any breach or non-compliance with any provision of this agreement.',
        '',
        '7. -Miscellaneous: This Agreement constitutes the entire understanding of the parties and is intended as a final expression of their agreement and cannot be altered, modified, amended or waived, in whole or in part, except by written instrument (email being sufficient) signed by both parties hereto. This agreement supersedes all prior agreements between the parties, whether oral or written. Should any provision of this agreement be held to be void, invalid or inoperative, such decision shall not affect any other provision hereof, and the remainder of this agreement shall be effective as though such void, invalid or inoperative provision had not been contained herein. No failure by Licensor hereto to perform any of its obligations hereunder shall be deemed a material breach of this agreement until the Licensee gives Licensor written notice of its failure to perform, and such failure has not been corrected within thirty (30) days from and after the service of such notice, or, if such breach is not reasonably capable of being cured within such thirty (30) day period, Licensor does not commence to cure such breach within said time period, and proceed with reasonable diligence to complete the curing of such breach thereafter. This agreement shall be governed by and interpreted in accordance with the laws of the Lagos,Nigeria applicable to agreements entered into and wholly performed in said State, without regard to any conflict of laws principles. You hereby agree that the exclusive jurisdiction and venue for any action, suit or proceeding based upon any matter, claim or controversy arising hereunder or relating hereto shall be in the state or federal courts located in the Lagos,Nigeria. You shall not be entitled to any monies in connection with the Master(s) other than as specifically set forth herein. All notices pursuant to this agreement shall be in writing and shall be given by registered or certified mail, return receipt requested (prepaid) at the respective addresses hereinabove set forth or such other address or addresses as may be designated by either party. Such notices shall be deemed given when received. Any notice mailed will be deemed to have been received five (5) business days after it is mailed; any notice dispatched by expedited delivery service will be deemed to be received two (2) business days after it is dispatched. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE READ THIS AGREEMENT AND HAVE BEEN ADVISED BY US OF THE SIGNIFICANT IMPORTANCE OF RETAINING AN INDEPENDENT ATTORNEY OF YOUR CHOICE TO REVIEW THIS AGREEMENT ON YOUR BEHALF. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE HAD THE UNRESTRICTED OPPORTUNITY TO BE REPRESENTED BY AN INDEPENDENT ATTORNEY. IN THE EVENT OF YOUR FAILURE TO OBTAIN AN INDEPENDENT ATTORNEY OR WAIVER THEREOF, YOU HEREBY WARRANT AND REPRESENT THAT YOU WILL NOT ATTEMPT TO USE SUCH FAILURE AND/OR WAIVER as a basis to avoid any obligations under this agreement, or to invalidate this agreement or To render this agreement or any part thereof unenforceable. This agreement may be executed in counterparts, each of which shall be deemed an original, and said counterparts shall constitute one and the same instrument. In addition, a signed copy of this agreement transmitted by facsimile or scanned into an image file and transmitted via email shall, for all purposes, be treated as if it was delivered containing an original manual signature of the party whose signature appears thereon and shall be binding upon such party as though an originally signed document had been delivered. Notwithstanding the foregoing, in the event that you do not sign this Agreement, your acknowledgment that you have reviewed the terms and conditions of this Agreement and your payment of the License Fee shall serve as your signature and acceptance of the terms and conditions of this Agreement.',
        
      ].join('\n')
    };
  }

  if (key.indexOf('wav') >= 0) {
    return {
      title: 'WAV License',
      content: [
        'This Non-Exclusive WAV License (Contract Preview Only) License Agreement (the "Agreement"), having been made on and effective as of [Effective Date] by and between Preview Only p/k/a Preview Only (the “Producer” or “Licensor”) and Licensee, sets forth the terms and conditions of Licensee’s use, and the rights granted in, the Producer’s instrumental music file entitled [Beat Title] (the "Beat") in consideration for Licensee’s payment of ₦77,000 (the "License Fee") on a so-called "WAV License (Contract Preview Only)" basis.',
        '',
        'This Agreement is issued solely in connection with and for Licensee’s use of the Beat pursuant and subject to all terms and conditions set forth herein.',
        '',
        '1. License Fee: The Licensee shall make payment of the License Fee to Licensor on the date of this Agreement. All rights granted to Licensee by Producer in the Beat are conditional upon Licensee’s timely payment of the License Fee. The License Fee is a one-time payment for the rights granted to Licensee and this Agreement is not valid until the License Fee has been paid.',
        '',
        '2. Delivery of the Beat:',
        '   - Licensor agrees to deliver the Beat as a high-quality WAV, MP3, as such terms are understood in the music industry.',
        '   - Licensor shall use commercially reasonable efforts to deliver the Beat to Licensee immediately after payment of the License Fee is made. Licensee will receive the Beat via email to the email address Licensee provided to Licensor.',
        '',
        '3. Term: The Term of this Agreement shall be ten (10) years and this license shall expire on the ten (10) year anniversary of the Effective Date.',
        '',
        '4. Use of the Beat:',
        '   - In consideration for Licensee’s payment of the License Fee, the Producer hereby grants Licensee a limited non-exclusive, nontransferable license and the right to incorporate, include and/or use the Beat in the preparation of one (1) new song or to incorporate the Beat into a new piece of instrumental music created by the Licensee. Licensee may create the new song or new instrumental music by recording his/her written lyrics over the Beat and/or by incorporating portions/samples of the Beat into pre-existing instrumental music written, produced and/or owned by Licensee. The new song or piece of instrumental music created by Licensee which incorporates some or all of the Beat shall be referred to as the "New Song". Permission is granted to Licensee to modify the arrangement, length, tempo, or pitch of the Beat in preparation of the New Song for public release.',
        '   - This License grants Licensee a worldwide, non-exclusive license to use the Beat as incorporated in the New Song in the manners and for the purposes expressly provided for herein, subject to the sale restrictions, limitations and prohibited uses stated in this Agreement. Licensee acknowledges and agrees that any and all rights granted to Licensee in the Beat pursuant to this Agreement are on a NON-EXCLUSIVE basis and Producer shall continue to license the Beat upon the same or similar terms and conditions as this Agreement to other potential third-party licensees.',
        '',
        'Permitted uses:',
        '   - The New Song may be used for any promotional purposes, including but not limited to, a release in a single format, for inclusion in a mixtape or free compilation of music bundled together (EP or album), and/or promotional, non-monetized digital streaming.',
        '   - Licensee may perform the song publicly for-profit performances and for Unlimited non-profit performances, including but not limited to, at a live performance (i.e. concert, festival, nightclub etc.), on terrestrial or satellite radio, and/or on the internet via third-party streaming services (Spotify, YouTube, iTunes Radio etc.). The New Song may be played on 2 terrestrial or satellite radio stations.',
        '   - The Licensee may use the New Song in synchronization with One (1) audiovisual work no longer than five (5) minutes in length (a "Video"). In the event that the New Song itself is longer than five (5) minutes in length, the Video may not play for longer than the length of the New Song. The Video may be broadcast on any television network and/or uploaded to the internet for digital streaming and/or free download by the public including but not limited to on YouTube and/or Vevo. Producer grants no other synchronization rights to Licensee.',
        '   - The Licensee may make the New Song available for sale in physical and/or digital form and sell 2,000 downloads/physical music products and are allowed 100,000 monetized audio streams, 1 monetized video streams, and are allowed unlimited free downloads. The New Song may be available for sale as a single and/or included in a compilation of other songs bundled together by Licensee as an EP or a full-length Album. The New Song may be sold via digital retailers for permanent digital download in mp3 format and/or physical format, including compact disc and vinyl records. For clarity and avoidance of doubt, the Licensee does NOT have the right to sell the Beat in the form that it was delivered to Licensee. The Licensee must create a New Song (or instrumental as detailed above) for its rights under this provision to vest. Any sale of the Beat in its original form by Licensee shall be a material breach of this Agreement and the Licensee shall be liable to the Licensor for damages as provided hereunder.',
        '',
        '5. Subject to the Licensee’s compliance with the terms and conditions of this Agreement, Licensee shall not be required to account or pay to Producer any royalties, fees, or monies paid to or collected by the Licensee (expressly excluding mechanical royalties), or which would otherwise be payable to Producer in connection with the use/exploitation of the New Song as set forth in this Agreement.',
        '',
        '6. Restrictions on the Use of the Beat:',
        '   - The rights granted to Licensee are NON-TRANSFERABLE and Licensee may not transfer or assign any of its rights hereunder to any third-party.',
        '   - The Licensee shall not synchronize, or permit third parties to synchronize, the Beat or New Song with any audiovisual works EXCEPT as expressly provided for and pursuant to Paragraph 4(b)(iii) of this Agreement for use in one (1) Video. This restriction includes, but is not limited to, use of the Beat and/or New Song in television, commercials, film/movies, theatrical works, video games, and in any other form on the Internet which is not expressly permitted herein.',
        '   - The Licensee shall not have the right to license or sublicense any use of the Beat or of the New Song, in whole or in part, for any so-called "samples".',
        '   - Licensee shall not engage in any unlawful copying, streaming, duplicating, selling, lending, renting, hiring, broadcasting, uploading, or downloading to any database, servers, computers, peer to peer sharing, or other file-sharing services, posting on websites, or distribution of the Beat in the form, or a substantially similar form, as delivered to Licensee. Licensee may send the Beat file to any individual musician, engineer, studio manager or other people who are working on the New Song.',
        '   - THE LICENSEE IS EXPRESSLY PROHIBITED FROM REGISTERING THE BEAT AND/OR NEW SONG WITH ANY CONTENT IDENTIFICATION SYSTEM, SERVICE PROVIDER, MUSIC DISTRIBUTOR, RECORD LABEL OR DIGITAL AGGREGATOR (for example TuneCore or CDBaby, and any other provider of user-generated content identification services). The purpose of this restriction is to prevent you from receiving a copyright infringement takedown notice from a third party who also received a non-exclusive license to use the Beat in a New Song. The Beat has already been tagged for Content Identification by Producer as a pre-emptive measure to protect all interested parties in the New Song. If you do not adhere to this policy, you are in violation of the terms of this License and your license to use the Beat and/or New Song may be revoked without notice or compensation to you.',
        '   - As applicable to both the underlying composition in the Beat and to the master recording of the Beat: (i) The parties acknowledge and agree that the New Song is a "derivative work", as that term is used in the United States Copyright Act; (ii) As applicable to the Beat and/or the New Song, there is no intention by the parties to create a joint work; and (iii) There is no intention by the Licensor to grant any rights in and/or to any other derivative works that may have been created by other third-party licensees.',
        '',
        '7. Ownership:',
        '   - The Producer is and shall remain the sole owner and holder of all rights, title, and interest in the Beat, including all copyrights to and in the sound recording and the underlying musical compositions written and composed by Producer. Nothing contained herein shall constitute an assignment by Producer to Licensee of any of the foregoing rights. Licensee may not, under any circumstances, register or attempt to register the New Song and/or the Beat with the U.S. Copyright Office. The aforementioned right to register the New Song and/or the Beat shall be strictly limited to Producer. Licensee will, upon request, execute, acknowledge and deliver to Producer such additional documents as Producer may deem necessary to evidence and effectuate Producer’s rights hereunder, and Licensee hereby grants to Producer the right as attorney-in-fact to execute, acknowledge, deliver and record in the U.S. Copyright Office or elsewhere any and all such documents if Licensee shall fail to execute same within five (5) days after so requested by Producer.',
        '   - For the avoidance of doubt, you do not own the master or the sound recording rights in the New Song. You have been licensed the right to use the Beat in the New Song and to commercially exploit the New Song based on the terms and conditions of this Agreement.',
        '   - Notwithstanding the above, you do own the lyrics or other original musical components of the New Song that were written or composed solely by you.',
        'With respect to the publishing rights and ownership of the underlying composition embodied in the New Song, the Licensee, and the Producer hereby acknowledge and agree that the underlying composition shall be owned/split between them as follows:',
        '   - Writer Share Licensor Name (Licensor) 50% Licensee Name (Licensee) 50%',
        '   - Producer shall own, control, and administer Fifty Percent (50%) of the so-called "Publishers Share" of the underlying composition.',
        '   - In the event that Licensee wishes to register his/her interests and rights to the underlying composition of the New Song with their Performing Rights Organization ("PRO"), Licensee must simultaneously identify and register the Producers share and ownership interest in the composition to indicate that Producer wrote and owns 50% of the composition in the New Song and as the owner of 50% of the Publishers share of the New Song.',
        '1. - The licensee shall be deemed to have signed, affirmed and ratified its acceptance of the terms of this Agreement by virtue of its payment of the License Fee to Licensor and its electronic acceptance of its terms and conditions at the time Licensee made payment of the License Fee.',
        '',
        '2. -Mechanical License: If any selection or musical composition, or any portion thereof, recorded in the New Song hereunder is written or composed by Producer, in whole or in part, alone or in collaboration with others, or is owned or controlled, in whole or in part, directly or indirectly, by Producer or any person, firm, or corporation in which Producer has a direct or indirect interest, then such selection and/or musical composition shall be hereinafter referred to as a "Controlled Composition". Producer hereby agrees to issue or cause to be issued, as applicable, to Licensee, mechanical licenses in respect of each Controlled Composition, which are embodied on the New Song. For that license, on the United States and Canada sales, Licensee will pay mechanical royalties at one hundred percent (100%) of the minimum statutory rate, subject to no cap of that rate for albums and/or EPs. For license outside the United States and Canada, the mechanical royalty rate will be the rate prevailing on an industry-wide basis in the country concerned on the date that this agreement has been entered into.',
        '',
        '3. -Credit: Licensee shall have the right to use and permit others to use Producers approved name, approved likeness, and other approved identification and approved biographical material concerning the Producer solely for purposes of trade and otherwise without restriction solely in connection with the New Song recorded hereunder. Licensee shall use best efforts to have Producer credited as a "producer" and shall give Producer appropriate production and songwriting credit on all compact discs, record, music video, and digital labels or any other record configuration manufactured which is now known or created in the future that embodies the New Song created hereunder and on all cover liner notes, any records containing the New Song and on the front and/or back cover of any album listing the New Song and other musician credits. The licensee shall use its best efforts to ensure that Producer is properly credited and Licensee shall check all proofs for the accuracy of credits, and shall use its best efforts to cure any mistakes regarding Producers credit. In the event of any failure by Licensee to issue the credit to Producer, Licensee must use reasonable efforts to correct any such failure immediately and on a prospective basis. Such credit shall be in the substantial form: "Produced by Preview Only.',
        '',
        '4. -Licensors Option: Licensor shall have the option, at Licensors sole discretion, to terminate this License at any time within three (3) years of the date of this Agreement upon written notice to Licensee. In the event that Licensor exercises this option, Licensor shall pay to Licensee a sum equal to Two Hundred Percent (200%) of the License Fee paid by Licensee. Upon Licensors exercise of the option, Licensee must immediately remove the New Song from any and all digital and physical distribution channels and must immediately cease access to any streams and/or downloads of the New Song by the general public.',
        '',
        '5. -Breach by Licensee: The licensee shall have five (5) business days from its receipt of written notice by Producer and/or Producers authorized representative to cure any alleged breach of this Agreement by Licensee. Licensees failure to cure the alleged breach within five (5) business days shall result in Licensees default of its obligations, its breach of this Agreement, and at Producers sole discretion, the termination of Licensees rights hereunder. If Licensee engages in the commercial exploitation and/or sale of the Beat or New Song outside of the manner and amount expressly provided for in this Agreement, Licensee shall be liable to Producer for monetary damages in an amount equal to any and all monies paid, collected by, or received by Licensee, or any third party on its behalf, in connection with such unauthorized commercial exploitation of the Beat and/or New Song. Licensee recognizes and agrees that a breach or threatened breach of this Agreement by Licensee give rise to irreparable injury to Producer, which may not be adequately compensated by damages. Accordingly, in the event of a breach or threatened breach by the Licensee of the provisions of this Agreement, Producer may seek and shall be entitled to a temporary restraining order and a preliminary injunction restraining the Licensee from violating the provisions of this Agreement. Nothing herein shall prohibit Producer from pursuing any other available legal or equitable remedy from such breach or threatened breach, including but not limited to the recovery of damages from Licensee. The Licensee shall be responsible for all costs, expenses or damages that Producer incurs as a result of any violation by the Licensee of any provision of this Agreement. Licensee obligation shall include court costs, litigation expenses, and reasonable attorneys fees.',
        '',
        'Warranties, Representations, and Indemnification: Licensee hereby agrees that Licensor has not made any guarantees or promises that the Beat fits the particular creative use or musical purpose intended or desired by the Licensee. The Beat, its sound recording, and the underlying musical composition embodied therein are licensed to the Licensee "as is" without warranties of any kind or fitness for a particular purpose.',
        '',
        '1. Parties hereto shall indemnify and hold each other harmless from any and all third party claims, liabilities, costs, losses, damages or expenses as are actually incurred by the non-defaulting party and shall hold the non-defaulting party, free, safe, and harmless against and from any and all claims, suits, demands, costs, liabilities, loss, damages, judgments, recoveries, costs, and expenses; (including, without limitation, reasonable attorneys fees), which may be made or brought, paid, or incurred by reason of any breach or claim of breach of the warranties and representations hereunder by the defaulting party, their agents, heirs, successors, assigns and employees, which have been reduced to final judgment; provided that prior to final judgment, arising out of any breach of any representations or warranties of the defaulting party contained in this agreement or any failure by defaulting party to perform any obligations on its part to be performed hereunder the non-defaulting party has given the defaulting party prompt written notice of all claims and the right to participate in the defense with counsel of its choice at its sole expense. In no event shall artist be entitled to seek injunctive or any other equitable relief for any breach or non-compliance with any provision of this agreement.',
        '',
        '7. -Miscellaneous: This Agreement constitutes the entire understanding of the parties and is intended as a final expression of their agreement and cannot be altered, modified, amended or waived, in whole or in part, except by written instrument (email being sufficient) signed by both parties hereto. This agreement supersedes all prior agreements between the parties, whether oral or written. Should any provision of this agreement be held to be void, invalid or inoperative, such decision shall not affect any other provision hereof, and the remainder of this agreement shall be effective as though such void, invalid or inoperative provision had not been contained herein. No failure by Licensor hereto to perform any of its obligations hereunder shall be deemed a material breach of this agreement until the Licensee gives Licensor written notice of its failure to perform, and such failure has not been corrected within thirty (30) days from and after the service of such notice, or, if such breach is not reasonably capable of being cured within such thirty (30) day period, Licensor does not commence to cure such breach within said time period, and proceed with reasonable diligence to complete the curing of such breach thereafter. This agreement shall be governed by and interpreted in accordance with the laws of the Lagos,Nigeria applicable to agreements entered into and wholly performed in said State, without regard to any conflict of laws principles. You hereby agree that the exclusive jurisdiction and venue for any action, suit or proceeding based upon any matter, claim or controversy arising hereunder or relating hereto shall be in the state or federal courts located in the Lagos,Nigeria. You shall not be entitled to any monies in connection with the Master(s) other than as specifically set forth herein. All notices pursuant to this agreement shall be in writing and shall be given by registered or certified mail, re [truncated]',
        

      ].join('\n')
    };
  }

  if (key.indexOf('premium') >= 0) {
    return {
      title: 'Premium Plus License',
      content: [
        'This Non-Exclusive Premium Plus License (Contract Preview Only) License Agreement (the "Agreement"), having been made on and effective as of [Effective Date] by and between Preview Only p/k/a Preview Only (the “Producer” or “Licensor”) and Licensee, sets forth the terms and conditions of Licensee’s use, and the rights granted in, the Producer’s instrumental music file entitled [Beat Title] (the "Beat") in consideration for Licensee’s payment of ₦155,000 (the "License Fee") on a so-called "Premium Plus License (Contract Preview Only)" basis.',
        '',
        'This Agreement is issued solely in connection with and for Licensee’s use of the Beat pursuant and subject to all terms and conditions set forth herein.',
        '',
        '1. License Fee: The Licensee shall make payment of the License Fee to Licensor on the date of this Agreement. All rights granted to Licensee by Producer in the Beat are conditional upon Licensee’s timely payment of the License Fee. The License Fee is a one-time payment for the rights granted to Licensee and this Agreement is not valid until the License Fee has been paid.',
        '',
        '2. Delivery of the Beat:',
        '   - Licensor agrees to deliver the Beat as a high-quality WAV, MP3, Track Stems, as such terms are understood in the music industry.',
        '   - Licensor shall use commercially reasonable efforts to deliver the Beat to Licensee immediately after payment of the License Fee is made. Licensee will receive the Beat via email to the email address Licensee provided to Licensor.',
        '',
        '3. Term: The Term of this Agreement shall be ten (10) years and this license shall expire on the ten (10) year anniversary of the Effective Date.',
        '',
        '4. Use of the Beat:',
        '   - In consideration for Licensee’s payment of the License Fee, the Producer hereby grants Licensee a limited non-exclusive, nontransferable license and the right to incorporate, include and/or use the Beat in the preparation of one (1) new song or to incorporate the Beat into a new piece of instrumental music created by the Licensee. Licensee may create the new song or new instrumental music by recording his/her written lyrics over the Beat and/or by incorporating portions/samples of the Beat into pre-existing instrumental music written, produced and/or owned by Licensee. The new song or piece of instrumental music created by Licensee which incorporates some or all of the Beat shall be referred to as the "New Song". Permission is granted to Licensee to modify the arrangement, length, tempo, or pitch of the Beat in preparation of the New Song for public release.',
        '   - This License grants Licensee a worldwide, non-exclusive license to use the Beat as incorporated in the New Song in the manners and for the purposes expressly provided for herein, subject to the sale restrictions, limitations and prohibited uses stated in this Agreement. Licensee acknowledges and agrees that any and all rights granted to Licensee in the Beat pursuant to this Agreement are on a NON-EXCLUSIVE basis and Producer shall continue to license the Beat upon the same or similar terms and conditions as this Agreement to other potential third-party licensees.',
        '',
        'Permitted uses:',
        '   - The New Song may be used for any promotional purposes, including but not limited to, a release in a single format, for inclusion in a mixtape or free compilation of music bundled together (EP or album), and/or promotional, non-monetized digital streaming.',
        '   - Licensee may perform the song publicly for-profit performances and for unlimited non-profit performances, including but not limited to, at a live performance (i.e. concert, festival, nightclub etc.), on terrestrial or satellite radio, and/or on the internet via third-party streaming services (Spotify, YouTube, iTunes Radio etc.). The New Song may be played on 2 terrestrial or satellite radio stations.',
        '   - The Licensee may use the New Song in synchronization with One (1) audiovisual work no longer than five (5) minutes in length (a "Video"). In the event that the New Song itself is longer than five (5) minutes in length, the Video may not play for longer than the length of the New Song. The Video may be broadcast on any television network and/or uploaded to the internet for digital streaming and/or free download by the public including but not limited to on YouTube and/or Vevo. Producer grants no other synchronization rights to Licensee.',
        '   - The Licensee may make the New Song available for sale in physical and/or digital form and sell 500,000 downloads/physical music products and are allowed 500,000 monetized audio streams, 1 monetized video streams, and are allowed unlimited free downloads. The New Song may be available for sale as a single and/or included in a compilation of other songs bundled together by Licensee as an EP or a full-length Album. The New Song may be sold via digital retailers for permanent digital download in mp3 format and/or physical format, including compact disc and vinyl records. For clarity and avoidance of doubt, the Licensee does NOT have the right to sell the Beat in the form that it was delivered to Licensee. The Licensee must create a New Song (or instrumental as detailed above) for its rights under this provision to vest. Any sale of the Beat in its original form by Licensee shall be a material breach of this Agreement and the Licensee shall be liable to the Licensor for damages as provided hereunder.',
        '',
        '5. Subject to the Licensee’s compliance with the terms and conditions of this Agreement, Licensee shall not be required to account or pay to Producer any royalties, fees, or monies paid to or collected by the Licensee (expressly excluding mechanical royalties), or which would otherwise be payable to Producer in connection with the use/exploitation of the New Song as set forth in this Agreement.',
        '',
        '6. Restrictions on the Use of the Beat:',
        '   - The rights granted to Licensee are NON-TRANSFERABLE and Licensee may not transfer or assign any of its rights hereunder to any third-party.',
        '   - The Licensee shall not synchronize, or permit third parties to synchronize, the Beat or New Song with any audiovisual works EXCEPT as expressly provided for and pursuant to Paragraph 4(b)(iii) of this Agreement for use in one (1) Video. This restriction includes, but is not limited to, use of the Beat and/or New Song in television, commercials, film/movies, theatrical works, video games, and in any other form on the Internet which is not expressly permitted herein.',
        '   - The Licensee shall not have the right to license or sublicense any use of the Beat or of the New Song, in whole or in part, for any so-called "samples".',
        '   - Licensee shall not engage in any unlawful copying, streaming, duplicating, selling, lending, renting, hiring, broadcasting, uploading, or downloading to any database, servers, computers, peer to peer sharing, or other file-sharing services, posting on websites, or distribution of the Beat in the form, or a substantially similar form, as delivered to Licensee. Licensee may send the Beat file to any individual musician, engineer, studio manager or other people who are working on the New Song.',
        '   - THE LICENSEE IS EXPRESSLY PROHIBITED FROM REGISTERING THE BEAT AND/OR NEW SONG WITH ANY CONTENT IDENTIFICATION SYSTEM, SERVICE PROVIDER, MUSIC DISTRIBUTOR, RECORD LABEL OR DIGITAL AGGREGATOR (for example TuneCore or CDBaby, and any other provider of user-generated content identification services). The purpose of this restriction is to prevent you from receiving a copyright infringement takedown notice from a third party who also received a non-exclusive license to use the Beat in a New Song. The Beat has already been tagged for Content Identification by Producer as a pre-emptive measure to protect all interested parties in the New Song. If you do not adhere to this policy, you are in violation of the terms of this License and your license to use the Beat and/or New Song may be revoked without notice or compensation to you.',
        '   - As applicable to both the underlying composition in the Beat and to the master recording of the Beat: (i) The parties acknowledge and agree that the New Song is a "derivative work", as that term is used in the United States Copyright Act; (ii) As applicable to the Beat and/or the New Song, there is no intention by the parties to create a joint work; and (iii) There is no intention by the Licensor to grant any rights in and/or to any other derivative works that may have been created by other third-party licensees.',
        '',
        '7. Ownership:',
        '   - The Producer is and shall remain the sole owner and holder of all rights, title, and interest in the Beat, including all copyrights to and in the sound recording and the underlying musical compositions written and composed by Producer. Nothing contained herein shall constitute an assignment by Producer to Licensee of any of the foregoing rights. Licensee may not, under any circumstances, register or attempt to register the New Song and/or the Beat with the U.S. Copyright Office. The aforementioned right to register the New Song and/or the Beat shall be strictly limited to Producer. Licensee will, upon request, execute, acknowledge and deliver to Producer such additional documents as Producer may deem necessary to evidence and effectuate Producer’s rights hereunder, and Licensee hereby grants to Producer the right as attorney-in-fact to execute, acknowledge, deliver and record in the U.S. Copyright Office or elsewhere any and all such documents if Licensee shall fail to execute same within five (5) days after so requested by Producer.',
        '   - For the avoidance of doubt, you do not own the master or the sound recording rights in the New Song. You have been licensed the right to use the Beat in the New Song and to commercially exploit the New Song based on the terms and conditions of this Agreement.',
        '   - Notwithstanding the above, you do own the lyrics or other original musical components of the New Song that were written or composed solely by you.',
        'With respect to the publishing rights and ownership of the underlying composition embodied in the New Song, the Licensee, and the Producer hereby acknowledge and agree that the underlying composition shall be owned/split between them as follows:',
        '',
        '   - Writer Share Licensor Name (Licensor) 50% Licensee Name (Licensee) 50%',
        '   - Producer shall own, control, and administer Fifty Percent (50%) of the so-called "Publishers Share" of the underlying composition.',
        '   - In the event that Licensee wishes to register his/her interests and rights to the underlying composition of the New Song with their Performing Rights Organization ("PRO"), Licensee must simultaneously identify and register the Producers share and ownership interest in the composition to indicate that Producer wrote and owns 50% of the composition in the New Song and as the owner of 50% of the Publishers share of the New Song.',
        '1. - The licensee shall be deemed to have signed, affirmed and ratified its acceptance of the terms of this Agreement by virtue of its payment of the License Fee to Licensor and its electronic acceptance of its terms and conditions at the time Licensee made payment of the License Fee.',
        '',
        '2. -Mechanical License: If any selection or musical composition, or any portion thereof, recorded in the New Song hereunder is written or composed by Producer, in whole or in part, alone or in collaboration with others, or is owned or controlled, in whole or in part, directly or indirectly, by Producer or any person, firm, or corporation in which Producer has a direct or indirect interest, then such selection and/or musical composition shall be hereinafter referred to as a "Controlled Composition". Producer hereby agrees to issue or cause to be issued, as applicable, to Licensee, mechanical licenses in respect of each Controlled Composition, which are embodied on the New Song. For that license, on the United States and Canada sales, Licensee will pay mechanical royalties at one hundred percent (100%) of the minimum statutory rate, subject to no cap of that rate for albums and/or EPs. For license outside the United States and Canada, the mechanical royalty rate will be the rate prevailing on an industry-wide basis in the country concerned on the date that this agreement has been entered into.',
        '',
        '3. -Credit: Licensee shall have the right to use and permit others to use Producers approved name, approved likeness, and other approved identification and approved biographical material concerning the Producer solely for purposes of trade and otherwise without restriction solely in connection with the New Song recorded hereunder. Licensee shall use best efforts to have Producer credited as a "producer" and shall give Producer appropriate production and songwriting credit on all compact discs, record, music video, and digital labels or any other record configuration manufactured which is now known or created in the future that embodies the New Song created hereunder and on all cover liner notes, any records containing the New Song and on the front and/or back cover of any album listing the New Song and other musician credits. The licensee shall use its best efforts to ensure that Producer is properly credited and Licensee shall check all proofs for the accuracy of credits, and shall use its best efforts to cure any mistakes regarding Producers credit. In the event of any failure by Licensee to issue the credit to Producer, Licensee must use reasonable efforts to correct any such failure immediately and on a prospective basis. Such credit shall be in the substantial form: "Produced by Preview Only.',
        '',
        '4. -Licensors Option: Licensor shall have the option, at Licensors sole discretion, to terminate this License at any time within three (3) years of the date of this Agreement upon written notice to Licensee. In the event that Licensor exercises this option, Licensor shall pay to Licensee a sum equal to Two Hundred Percent (200%) of the License Fee paid by Licensee. Upon Licensors exercise of the option, Licensee must immediately remove the New Song from any and all digital and physical distribution channels and must immediately cease access to any streams and/or downloads of the New Song by the general public.',
        '',
        '5. -Breach by Licensee: The licensee shall have five (5) business days from its receipt of written notice by Producer and/or Producers authorized representative to cure any alleged breach of this Agreement by Licensee. Licensees failure to cure the alleged breach within five (5) business days shall result in Licensees default of its obligations, its breach of this Agreement, and at Producers sole discretion, the termination of Licensees rights hereunder. If Licensee engages in the commercial exploitation and/or sale of the Beat or New Song outside of the manner and amount expressly provided for in this Agreement, Licensee shall be liable to Producer for monetary damages in an amount equal to any and all monies paid, collected by, or received by Licensee, or any third party on its behalf, in connection with such unauthorized commercial exploitation of the Beat and/or New Song. Licensee recognizes and agrees that a breach or threatened breach of this Agreement by Licensee give rise to irreparable injury to Producer, which may not be adequately compensated by damages. Accordingly, in the event of a breach or threatened breach by the Licensee of the provisions of this Agreement, Producer may seek and shall be entitled to a temporary restraining order and a preliminary injunction restraining the Licensee from violating the provisions of this Agreement. Nothing herein shall prohibit Producer from pursuing any other available legal or equitable remedy from such breach or threatened breach, including but not limited to the recovery of damages from Licensee. The Licensee shall be responsible for all costs, expenses or damages that Producer incurs as a result of any violation by the Licensee of any provision of this Agreement. Licensee obligation shall include court costs, litigation expenses, and reasonable attorneys fees.',
        '',
        'Warranties, Representations, and Indemnification: Licensee hereby agrees that Licensor has not made any guarantees or promises that the Beat fits the particular creative use or musical purpose intended or desired by the Licensee. The Beat, its sound recording, and the underlying musical composition embodied therein are licensed to the Licensee "as is" without warranties of any kind or fitness for a particular purpose.',
        '',
        '1. Parties hereto shall indemnify and hold each other harmless from any and all third party claims, liabilities, costs, losses, damages or expenses as are actually incurred by the non-defaulting party and shall hold the non-defaulting party, free, safe, and harmless against and from any and all claims, suits, demands, costs, liabilities, loss, damages, judgments, recoveries, costs, and expenses; (including, without limitation, reasonable attorneys fees), which may be made or brought, paid, or incurred by reason of any breach or claim of breach of the warranties and representations hereunder by the defaulting party, their agents, heirs, successors, assigns and employees, which have been reduced to final judgment; provided that prior to final judgment, arising out of any breach of any representations or warranties of the defaulting party contained in this agreement or any failure by defaulting party to perform any obligations on its part to be performed hereunder the non-defaulting party has given the defaulting party prompt written notice of all claims and the right to participate in the defense with counsel of its choice at its sole expense. In no event shall artist be entitled to seek injunctive or any other equitable relief for any breach or non-compliance with any provision of this agreement.',
        '',
        '7. -Miscellaneous: This Agreement constitutes the entire understanding of the parties and is intended as a final expression of their agreement and cannot be altered, modified, amended or waived, in whole or in part, except by written instrument (email being sufficient) signed by both parties hereto. This agreement supersedes all prior agreements between the parties, whether oral or written. Should any provision of this agreement be held to be void, invalid or inoperative, such decision shall not affect any other provision hereof, and the remainder of this agreement shall be effective as though such void, invalid or inoperative provision had not been contained herein. No failure by Licensor hereto to perform any of its obligations hereunder shall be deemed a material breach of this agreement until the Licensee gives Licensor written notice of its failure to perform, and such failure has not been corrected within thirty (30) days from and after the service of such notice, or, if such breach is not reasonably capable of being cured within such thirty (30) day period, Licensor does not commence to cure such breach within said time period, and proceed with reasonable diligence to complete the curing of such breach thereafter. This agreement shall be governed by and interpreted in accordance with the laws of the **Abuja,Nigeria** applicable to agreements entered into and wholly performed in said State, without regard to any conflict of laws principles. You hereby agree that the exclusive jurisdiction and venue for any action, suit or proceeding based upon any matter, claim or controversy arising hereunder or relating hereto shall be in the state or federal courts located in the **Abuja,Nigeria**. You shall not be entitled to any monies in connection with the Master(s) other than as specifically set forth herein. All notices pursuant to this agreement shall be in writing and shall be given by registered or certified mail, return receipt requested (prepaid) at the respective addresses hereinabove set forth or such other address or addresses as may be designated by either party. Such notices shall be deemed given when received. Any notice mailed will be deemed to have been received five (5) business days after it is mailed; any notice dispatched by expedited delivery service will be deemed to be received two (2) business days after it is dispatched. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE READ THIS AGREEMENT AND HAVE BEEN ADVISED BY US OF THE SIGNIFICANT IMPORTANCE OF RETAINING AN INDEPENDENT ATTORNEY OF YOUR CHOICE TO REVIEW THIS AGREEMENT ON YOUR BEHALF. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE HAD THE UNRESTRICTED OPPORTUNITY TO BE REPRESENTED BY AN INDEPENDENT ATTORNEY. IN THE EVENT OF YOUR FAILURE TO OBTAIN AN INDEPENDENT ATTORNEY OR WAIVER THEREOF, YOU HEREBY WARRANT AND REPRESENT THAT YOU WILL NOT ATTEMPT TO USE SUCH FAILURE AND/OR WAIVER as a basis to avoid any obligations under this agreement, or to invalidate this agreement or To render this agreement or any part thereof unenforceable. This agreement may be executed in counterparts, each of which shall be deemed an original, and said counterparts shall constitute one and the same instrument. In addition, a signed copy of this agreement transmitted by facsimile or scanned into an image file and transmitted via email shall, for all purposes, be treated as if it was delivered containing an original manual signature of the party whose signature appears thereon and shall be binding upon such party as though an originally signed document had been delivered. Notwithstanding the foregoing, in the event that you do not sign this Agreement, your acknowledgment that you have reviewed the terms and conditions of this Agreement and your payment of the License Fee shall serve as your signature and acceptance of the terms and conditions of this Agreement.'
      ].join('\n')
    };
  }

  if (key.indexOf('unlimited') >= 0) {
    return {
      title: 'Unlimited License',
      content: [
        'This Non-Exclusive Unlimited License (Contract Preview Only) License Agreement (the "Agreement"), having been made on and effective as of [Effective Date] by and between Preview Only p/k/a Preview Only (the “Producer” or “Licensor”) and Licensee, sets forth the terms and conditions of Licensee’s use, and the rights granted in, the Producer’s instrumental music file entitled [Beat Title] (the "Beat") in consideration for Licensee’s payment of ₦230,000 (the "License Fee") on a so-called "Unlimited License (Contract Preview Only)" basis.',
        '',
        'This Agreement is issued solely in connection with and for Licensee’s use of the Beat pursuant and subject to all terms and conditions set forth herein.',
        '',
        '1. License Fee: The Licensee shall make payment of the License Fee to Licensor on the date of this Agreement. All rights granted to Licensee by Producer in the Beat are conditional upon Licensee’s timely payment of the License Fee. The License Fee is a one-time payment for the rights granted to Licensee and this Agreement is not valid until the License Fee has been paid.',
        '',
        '2. Delivery of the Beat:',
        '   - Licensor agrees to deliver the Beat as a high-quality WAV, MP3, Track Stems, as such terms are understood in the music industry.',
        '   - Licensor shall use commercially reasonable efforts to deliver the Beat to Licensee immediately after payment of the License Fee is made. Licensee will receive the Beat via email to the email address Licensee provided to Licensor.',
        '',
        '3. Term: The Term of this Agreement shall be unlimited and this license shall not expire.',
        '',
        '4. Use of the Beat:',
        '   - In consideration for Licensee’s payment of the License Fee, the Producer hereby grants Licensee a limited non-exclusive, nontransferable license and the right to incorporate, include and/or use the Beat in the preparation of one (1) new song or to incorporate the Beat into a new piece of instrumental music created by the Licensee. Licensee may create the new song or new instrumental music by recording his/her written lyrics over the Beat and/or by incorporating portions/samples of the Beat into pre-existing instrumental music written, produced and/or owned by Licensee. The new song or piece of instrumental music created by Licensee which incorporates some or all of the Beat shall be referred to as the "New Song". Permission is granted to Licensee to modify the arrangement, length, tempo, or pitch of the Beat in preparation of the New Song for public release.',
        '   - This License grants Licensee a worldwide, non-exclusive license to use the Beat as incorporated in the New Song in the manners and for the purposes expressly provided for herein, subject to the sale restrictions, limitations and prohibited uses stated in this Agreement. Licensee acknowledges and agrees that any and all rights granted to Licensee in the Beat pursuant to this Agreement are on a NON-EXCLUSIVE basis and Producer shall continue to license the Beat upon the same or similar terms and conditions as this Agreement to other potential third-party licensees.',
        '',
        'Permitted uses:',
        '   - The New Song may be used for any promotional purposes, including but not limited to, a release in a single format, for inclusion in a mixtape or free compilation of music bundled together (EP or album), and/or promotional, non-monetized digital streaming.',
        '   - Licensee may perform the song publicly for-profit performances and for unlimited non-profit performances, including but not limited to, at a live performance (i.e. concert, festival, nightclub etc.), on terrestrial or satellite radio, and/or on the internet via third-party streaming services (Spotify, YouTube, iTunes Radio etc.). The New Song may be played on Unlimited terrestrial or satellite radio stations.',
        '   - The Licensee may use the New Song in synchronization with One (1) audiovisual work no longer than five (5) minutes in length (a "Video"). In the event that the New Song itself is longer than five (5) minutes in length, the Video may not play for longer than the length of the New Song. The Video may be broadcast on any television network and/or uploaded to the internet for digital streaming and/or free download by the public including but not limited to on YouTube and/or Vevo. Producer grants no other synchronization rights to Licensee.',
        '   - The Licensee may make the New Song available for sale in physical and/or digital form and sell unlimited downloads/physical music products and are allowed unlimited monetized audio streams, 1 monetized video streams, and are allowed unlimited free downloads. The New Song may be available for sale as a single and/or included in a compilation of other songs bundled together by Licensee as an EP or a full-length Album. The New Song may be sold via digital retailers for permanent digital download in mp3 format and/or physical format, including compact disc and vinyl records. For clarity and avoidance of doubt, the Licensee does NOT have the right to sell the Beat in the form that it was delivered to Licensee. The Licensee must create a New Song (or instrumental as detailed above) for its rights under this provision to vest. Any sale of the Beat in its original form by Licensee shall be a material breach of this Agreement and the Licensee shall be liable to the Licensor for damages as provided hereunder.',
        '',
        '5. Subject to the Licensee’s compliance with the terms and conditions of this Agreement, Licensee shall not be required to account or pay to Producer any royalties, fees, or monies paid to or collected by the Licensee (expressly excluding mechanical royalties), or which would otherwise be payable to Producer in connection with the use/exploitation of the New Song as set forth in this Agreement.',
        '',
        '6. Restrictions on the Use of the Beat:',
        '   - The rights granted to Licensee are NON-TRANSFERABLE and Licensee may not transfer or assign any of its rights hereunder to any third-party.',
        '   - The Licensee shall not synchronize, or permit third parties to synchronize, the Beat or New Song with any audiovisual works EXCEPT as expressly provided for and pursuant to Paragraph 4(b)(iii) of this Agreement for use in one (1) Video. This restriction includes, but is not limited to, use of the Beat and/or New Song in television, commercials, film/movies, theatrical works, video games, and in any other form on the Internet which is not expressly permitted herein.',
        '   - The Licensee shall not have the right to license or sublicense any use of the Beat or of the New Song, in whole or in part, for any so-called "samples".',
        '   - Licensee shall not engage in any unlawful copying, streaming, duplicating, selling, lending, renting, hiring, broadcasting, uploading, or downloading to any database, servers, computers, peer to peer sharing, or other file-sharing services, posting on websites, or distribution of the Beat in the form, or a substantially similar form, as delivered to Licensee. Licensee may send the Beat file to any individual musician, engineer, studio manager or other people who are working on the New Song.',
        '   - THE LICENSEE IS EXPRESSLY PROHIBITED FROM REGISTERING THE BEAT AND/OR NEW SONG WITH ANY CONTENT IDENTIFICATION SYSTEM, SERVICE PROVIDER, MUSIC DISTRIBUTOR, RECORD LABEL OR DIGITAL AGGREGATOR (for example TuneCore or CDBaby, and any other provider of user-generated content identification services). The purpose of this restriction is to prevent you from receiving a copyright infringement takedown notice from a third party who also received a non-exclusive license to use the Beat in a New Song. The Beat has already been tagged for Content Identification by Producer as a pre-emptive measure to protect all interested parties in the New Song. If you do not adhere to this policy, you are in violation of the terms of this License and your license to use the Beat and/or New Song may be revoked without notice or compensation to you.',
        '   - As applicable to both the underlying composition in the Beat and to the master recording of the Beat: (i) The parties acknowledge and agree that the New Song is a "derivative work", as that term is used in the United States Copyright Act; (ii) As applicable to the Beat and/or the New Song, there is no intention by the parties to create a joint work; and (iii) There is no intention by the Licensor to grant any rights in and/or to any other derivative works that may have been created by other third-party licensees.',
        '',
        '7. Ownership:',
        '   - The Producer is and shall remain the sole owner and holder of all rights, title, and interest in the Beat, including all copyrights to and in the sound recording and the underlying musical compositions written and composed by Producer. Nothing contained herein shall constitute an assignment by Producer to Licensee of any of the foregoing rights. Licensee may not, under any circumstances, register or attempt to register the New Song and/or the Beat with the U.S. Copyright Office. The aforementioned right to register the New Song and/or the Beat shall be strictly limited to Producer. Licensee will, upon request, execute, acknowledge and deliver to Producer such additional documents as Producer may deem necessary to evidence and effectuate Producer’s rights hereunder, and Licensee hereby grants to Producer the right as attorney-in-fact to execute, acknowledge, deliver and record in the U.S. Copyright Office or elsewhere any and all such documents if Licensee shall fail to execute same within five (5) days after so requested by Producer.',
        '   - For the avoidance of doubt, you do not own the master or the sound recording rights in the New Song. You have been licensed the right to use the Beat in the New Song and to commercially exploit the New Song based on the terms and conditions of this Agreement.',
        '   - Notwithstanding the above, you do own the lyrics or other original musical components of the New Song that were written or composed solely by you.',
        '',
        'With respect to the publishing rights and ownership of the underlying composition embodied in the New Song, the Licensee, and the Producer hereby acknowledge and agree that the underlying composition shall be owned/split between them as follows:',
        '',
        '   - Writer Share Licensor Name (Licensor) 50% Licensee Name (Licensee) 50%',
        '   - Producer shall own, control, and administer Fifty Percent (50%) of the so-called "Publishers Share" of the underlying composition.',
        '   - In the event that Licensee wishes to register his/her interests and rights to the underlying composition of the New Song with their Performing Rights Organization ("PRO"), Licensee must simultaneously identify and register the Producers share and ownership interest in the composition to indicate that Producer wrote and owns 50% of the composition in the New Song and as the owner of 50% of the Publishers share of the New Song.',
        '1. - The licensee shall be deemed to have signed, affirmed and ratified its acceptance of the terms of this Agreement by virtue of its payment of the License Fee to Licensor and its electronic acceptance of its terms and conditions at the time Licensee made payment of the License Fee.',
        '',
        '2. -Mechanical License: If any selection or musical composition, or any portion thereof, recorded in the New Song hereunder is written or composed by Producer, in whole or in part, alone or in collaboration with others, or is owned or controlled, in whole or in part, directly or indirectly, by Producer or any person, firm, or corporation in which Producer has a direct or indirect interest, then such selection and/or musical composition shall be hereinafter referred to as a "Controlled Composition". Producer hereby agrees to issue or cause to be issued, as applicable, to Licensee, mechanical licenses in respect of each Controlled Composition, which are embodied on the New Song. For that license, on the United States and Canada sales, Licensee will pay mechanical royalties at one hundred percent (100%) of the minimum statutory rate, subject to no cap of that rate for albums and/or EPs. For license outside the United States and Canada, the mechanical royalty rate will be the rate prevailing on an industry-wide basis in the country concerned on the date that this agreement has been entered into.',
        '',
        '3. -Credit: Licensee shall have the right to use and permit others to use Producers approved name, approved likeness, and other approved identification and approved biographical material concerning the Producer solely for purposes of trade and otherwise without restriction solely in connection with the New Song recorded hereunder. Licensee shall use best efforts to have Producer credited as a "producer" and shall give Producer appropriate production and songwriting credit on all compact discs, record, music video, and digital labels or any other record configuration manufactured which is now known or created in the future that embodies the New Song created hereunder and on all cover liner notes, any records containing the New Song and on the front and/or back cover of any album listing the New Song and other musician credits. The licensee shall use its best efforts to ensure that Producer is properly credited and Licensee shall check all proofs for the accuracy of credits, and shall use its best efforts to cure any mistakes regarding Producers credit. In the event of any failure by Licensee to issue the credit to Producer, Licensee must use reasonable efforts to correct any such failure immediately and on a prospective basis. Such credit shall be in the substantial form: "Produced by Preview Only.',
        '',
        '4. -Licensors Option: Licensor shall have the option, at Licensors sole discretion, to terminate this License at any time within three (3) years of the date of this Agreement upon written notice to Licensee. In the event that Licensor exercises this option, Licensor shall pay to Licensee a sum equal to Two Hundred Percent (200%) of the License Fee paid by Licensee. Upon Licensors exercise of the option, Licensee must immediately remove the New Song from any and all digital and physical distribution channels and must immediately cease access to any streams and/or downloads of the New Song by the general public.',
        '',
        '5. -Breach by Licensee: The licensee shall have five (5) business days from its receipt of written notice by Producer and/or Producers authorized representative to cure any alleged breach of this Agreement by Licensee. Licensees failure to cure the alleged breach within five (5) business days shall result in Licensees default of its obligations, its breach of this Agreement, and at Producers sole discretion, the termination of Licensees rights hereunder. If Licensee engages in the commercial exploitation and/or sale of the Beat or New Song outside of the manner and amount expressly provided for in this Agreement, Licensee shall be liable to Producer for monetary damages in an amount equal to any and all monies paid, collected by, or received by Licensee, or any third party on its behalf, in connection with such unauthorized commercial exploitation of the Beat and/or New Song. Licensee recognizes and agrees that a breach or threatened breach of this Agreement by Licensee give rise to irreparable injury to Producer, which may not be adequately compensated by damages. Accordingly, in the event of a breach or threatened breach by the Licensee of the provisions of this Agreement, Producer may seek and shall be entitled to a temporary restraining order and a preliminary injunction restraining the Licensee from violating the provisions of this Agreement. Nothing herein shall prohibit Producer from pursuing any other available legal or equitable remedy from such breach or threatened breach, including but not limited to the recovery of damages from Licensee. The Licensee shall be responsible for all costs, expenses or damages that Producer incurs as a result of any violation by the Licensee of any provision of this Agreement. Licensee obligation shall include court costs, litigation expenses, and reasonable attorneys fees.',
        '',
        'Warranties, Representations, and Indemnification: Licensee hereby agrees that Licensor has not made any guarantees or promises that the Beat fits the particular creative use or musical purpose intended or desired by the Licensee. The Beat, its sound recording, and the underlying musical composition embodied therein are licensed to the Licensee "as is" without warranties of any kind or fitness for a particular purpose.',
        '',
        '1. Parties hereto shall indemnify and hold each other harmless from any and all third party claims, liabilities, costs, losses, damages or expenses as are actually incurred by the non-defaulting party and shall hold the non-defaulting party, free, safe, and harmless against and from any and all claims, suits, demands, costs, liabilities, loss, damages, judgments, recoveries, costs, and expenses; (including, without limitation, reasonable attorneys fees), which may be made or brought, paid, or incurred by reason of any breach or claim of breach of the warranties and representations hereunder by the defaulting party, their agents, heirs, successors, assigns and employees, which have been reduced to final judgment; provided that prior to final judgment, arising out of any breach of any representations or warranties of the defaulting party contained in this agreement or any failure by defaulting party to perform any obligations on its part to be performed hereunder the non-defaulting party has given the defaulting party prompt written notice of all claims and the right to participate in the defense with counsel of its choice at its sole expense. In no event shall artist be entitled to seek injunctive or any other equitable relief for any breach or non-compliance with any provision of this agreement.',
        '',
        'Miscellaneous: This Agreement constitutes the entire understanding of the parties and is intended as a final expression of their agreement and cannot be altered, modified, amended or waived, in whole or in part, except by written instrument (email being sufficient) signed by both parties hereto. This agreement supersedes all prior agreements between the parties, whether oral or written. Should any provision of this agreement be held to be void, invalid or inoperative, such decision shall not affect any other provision hereof, and the remainder of this agreement shall be effective as though such void, invalid or inoperative provision had not been contained herein. No failure by Licensor hereto to perform any of its obligations hereunder shall be deemed a material breach of this agreement until the Licensee gives Licensor written notice of its failure to perform, and such failure has not been corrected within thirty (30) days from and after the service of such notice, or, if such breach is not reasonably capable of being cured within such thirty (30) day period, Licensor does not commence to cure such breach within said time period, and proceed with reasonable diligence to complete the curing of such breach thereafter. This agreement shall be governed by and interpreted in accordance with the laws of the **Abuja,Nigeria** applicable to agreements entered into and wholly performed in said State, without regard to any conflict of laws principles. You hereby agree that the exclusive jurisdiction and venue for any action, suit or proceeding based upon any matter, claim or controversy arising hereunder or relating hereto shall be in the state or federal courts located in the **Abuja,Nigeria**. You shall not be entitled to any monies in connection with the Master(s) other than as specifically set forth herein. All notices pursuant to this agreement shall be in writing and shall be given by registered or certified mail, return receipt requested (prepaid) at the respective addresses hereinabove set forth or such other address or addresses as may be designated by either party. Such notices shall be deemed given when received. Any notice mailed will be deemed to have been received five (5) business days after it is mailed; any notice dispatched by expedited delivery service will be deemed to be received two (2) business days after it is dispatched. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE READ THIS AGREEMENT AND HAVE BEEN ADVISED BY US OF THE SIGNIFICANT IMPORTANCE OF RETAINING AN INDEPENDENT ATTORNEY OF YOUR CHOICE TO REVIEW THIS AGREEMENT ON YOUR BEHALF. YOU ACKNOWLEDGE AND AGREE THAT YOU HAVE HAD THE UNRESTRICTED OPPORTUNITY TO BE REPRESENTED BY AN INDEPENDENT ATTORNEY. IN THE EVENT OF YOUR FAILURE TO OBTAIN AN INDEPENDENT ATTORNEY OR WAIVER THEREOF, YOU HEREBY WARRANT AND REPRESENT THAT YOU WILL NOT ATTEMPT TO USE SUCH FAILURE AND/OR WAIVER as a basis to avoid any obligations under this agreement, or to invalidate this agreement or To render this agreement or any part thereof unenforceable.'
      ].join('\n')
    };
  }

  if (key.indexOf('exclusive') >= 0) {
    return {
      title: 'Exclusive License',
      content: [
        'Preview Only -w- Licensee / Producer Agreement / "Preview Track Only"',
        '',
        'The following sets forth the material terms and conditions with respect to Preview Only ("Producer", "me", "we", or the like) producing certain recording(s) embodying the musical performance of Licensee ("Artist", "you", "your", or the like). In the event the number of master recordings hereunder is no more than one (1), then all references to "Masters" hereunder shall be read and deemed to refer to one (1) "Master." For good and valuable consideration (the receipt and sufficiency of which is hereby acknowledged), the parties hereby agree as follows:',
        '',
        '1. Product Commitment: Producer shall produce one (1) musical composition entitled "Preview Track Only" (the "Composition") embodying Artist’s featured performance of a yet-to-be-titled master recording (the "Master") for delivery to Artist for, among other exploitations, the manufacture and distribution of records. For the avoidance of doubt, the Composition provided by Producer to create the Master shall be solely retained and owned by Producer as a pre-existing composition, and the composition made by Artist hereunder is a derivative. The territory of this agreement shall be the Universe. Artist acknowledges the satisfactory delivery, receipt, and acceptance of the Master.',
        '',
        '2. Rights: The Master (expressly excluding the underlying musical composition), from the inception of creation, shall be considered a "work made for hire" for Artist (or Artist’s designees) within the meaning of the Copyright Act of 1976 (Title 17, U.S.C.). If it is determined that any Master does not so qualify, then that Master, together with all rights therein (including the sound recording copyright(s) but excluding the underlying musical composition) shall hereby be deemed transferred to Artist. Subject to the terms and conditions contained in this agreement, Artist shall have the sole and exclusive right in perpetuity and throughout the universe, including, without limitation: (i) to manufacture, advertise, sell, license or otherwise dispose of the Master and derivatives derived therefrom in any manner or media whatsoever upon such terms, and under such trademarks, as Artist elects, or, in Artist’s sole discretion, to refrain therefrom; (ii) to perform the Master publicly and to permit the public performance thereof by any method now or hereafter known; and (iii) to include Producer’s audio performance in an audio-visual production ("Video"). Notwithstanding the foregoing, Artist (or its designees) shall have no right to make any edits/changes to Producer’s composition; no right to use Producer’s composition apart from the Master; and no right to use Producer’s composition in a way to imply any sort of endorsement.',
        '',
        '3. Fee: Artist shall pay to Producer a non-returnable, non-recoupable fee in the amount of ₦0.00 (the "Fee"). The Fee shall be payable upon the full execution of this agreement.',
        '',
        '4. Controlled Compositions: If any selection or musical composition, or any portion thereof, recorded in the Masters hereunder is written or composed by Producer, in whole or in part, alone or in collaboration with others, or is owned or controlled, in whole or in part, directly or indirectly, by Producer or any person, firm, or corporation in which Producer has a direct or indirect interest, then such selection and/or musical composition shall be hereinafter referred to as a "Controlled Composition". Producer hereby agrees to issue or cause to be issued, as applicable, to Artist, or Artist’s designees, mechanical licenses (including, without limitation, any "first use" mechanical licenses) and other licenses in respect of each Controlled Composition subject to the same terms and conditions as are applicable to selections and musical compositions written, composed, owned, or controlled, in whole or in part, directly or indirectly, by Artist, which are embodied on the Master or any recording recorded under this agreement. For that license, on the United States and Canada sales, Artist will pay mechanical royalties at one hundred percent (100%) of the minimum statutory rate, subject to no cap of that rate for albums and/or EPs. For license outside the United States and Canada, the mechanical royalty rate will be the rate prevailing on an industry-wide basis in the country concerned on the date that this agreement has been entered into. Producer hereby grants Artist a license to reproduce Controlled Compositions that are embodied on the Master produced hereunder in synchronization with and in time relation to visual images featuring Artist’s performances in so-called promotional "video programs", on a royalty-free basis and in so-called commercial "video programs", it being understood that Producer shall be entitled to Producer’s pro-rata share of any royalties attributable to such commercial "video programs". Upon Artist’s request Producer shall ...',
        '',
        '5. Publishing Rights: With respect to the publishing rights and ownership of the underlying composition embodied in the Master, the Artist, and the Producer hereby acknowledge and agree that the underlying composition shall be owned/split between them as follows:',
        '   - LICENSEE OWNS 50% OF PUBLISHING RIGHTS While Producer owns 50%',
        '',
        '6. Credit and Likeness: Artist shall have the right to use and permit others to use Producer’s approved name, approved likeness, and other approved identification and approved biographical material concerning the Producer for purposes of trade and otherwise without restriction solely in connection with the Masters recorded hereunder. Artist shall accord (or shall cause to accord) Producer a credit on (i) labels and liner notes of the Master, where applicable, (ii) in all configurations (including in applicable meta-data) derived from the Master, (iii) in all trade and consumer advertisements, including Billboard Magazine strip ads, which pertain exclusively to the Masters hereunder, that are one-quarter (1/4) page or larger in size, placed directly by Artist, and appear in so-called "nationwide" trade publications in the United States. Artist shall ensure that Producer is properly credited and Artist shall check all proofs for the accuracy of credits, and shall cure any mistakes regarding Producer’s credit. Such credit shall be in the substantial form: "Produced by Preview Only".',
        '',
        '7. Warranties, Representations, and Indemnification:',
        '   - Artist hereby agrees that Producer has not made any guarantees or promises that the Master fits the particular creative use or musical purpose intended or desired by the Artist. The Master, its sound recording, and the Composition embodied therein are delivered to the Artist "as is" without warranties of any kind or fitness for a particular purpose. Artist further acknowledges and agrees that the Composition produced by Producer may previously have been licensed to third parties on a non-exclusive basis. Any licenses granted by Producer, which precede this agreement, shall remain in effect and shall not be affected by this agreement. Artist waives any claims against Producer for any pre-existing licenses for the Composition. Artist further agrees that Artist will not submit any claims against the third-party licensees for their non-exclusive use of the Composition.',
        '   - Parties hereto shall indemnify and hold each other harmless from any and all third party claims, liabilities, costs, losses, damages or expenses as are actually incurred by the non-defaulting party and shall hold the non-defaulting party, free, safe, and harmless against and from any and all claims, suits, demands, costs, liabilities, loss, damages, judgments, recoveries, costs, and expenses; (including, without limitation, reasonable outside attorneys’ fees), which may be made or brought, paid, or incurred by reason of any breach or claim of breach of the warranties and representations hereunder by the defaulting party, their agents, heirs, successors, assigns and employees, which have been reduced to final judgment; provided that prior to final judgment, arising out of any breach of any representations or warranties of the defaulting party contained in this agreement or any failure by defaulting party to perform any obligations on its part to be performed hereunder the non-defaulting party has given the defaulting party prompt written notice of all claims and the right to participate in the defense with counsel of its choice at its sole expense. In no event shall Artist be entitled to seek injunctive or any other equitable relief for any breach or non-compliance with any provision of this agreement.',
        '',
        '8. Miscellaneous: This agreement has been entered into in the LAGOS STATE, Nigeria and the validity, interpretation, and legal effect of this agreement shall be governed by the laws of the LAGOS STATE, Nigeria applicable to contracts entered into and performed entirely within such State. The courts of LAGOS STATE, Nigeria (state and federal) only will have jurisdiction of any controversies regarding this agreement and the parties hereto consent to the jurisdiction of said courts. All notices, statements, and payments to be sent to any party hereunder shall be addressed to such party at the applicable address set forth on the first page hereof or at such other address as is designated in writing by the applicable party from time to time. All notices shall be in writing and shall either be served by personal delivery (with a written receipt of such delivery) or certified or registered mail, return receipt requested, all charges prepaid, except statements may be sent by regular U.S. mail. Except as otherwise provided herein, notices delivered in accordance with the foregoing shall be deemed given when personally delivered, or five (5) days after mailing, except that notices of change of address shall be effective only after actual receipt. Where approvals are required hereunder, such approval or consent shall not be unreasonably withheld and the parties acknowledge and agree that email confirmations/responses shall suffice. Producer shall not be entitled to any monies in connection with the Master(s) other than as specifically set forth herein. Producer shall have the right to assign this agreement to any parent, subsidiary, or affiliate, or any individual or entity owning or acquiring a substantial portion of Producer’s stock or assets provided that Producer remains secondary liable. Artist may not assign any of Artists rights or obligations hereunder without Producers prior written consent and any such purported assignment shall be null and void ab initio. Both parties agree and acknowledge that this agreement (a) will be binding upon and inure to the benefit of the parties hereto and their respective successors, permitted assigns, heirs, estates, administrators, and executors; (b) embodies the sole and entire agreement of the parties in respect of, and supersedes all prior oral or written understandings between them concerning the subject matter hereof; and (c) may not be amended except by a written instrument signed by all parties hereto. A waiver by either party hereto of any provision of this agreement in any instance shall not be deemed to be a waiver for the future. All remedies, rights, undertakings, and obligations contained in this agreement shall be cumulative and none of them shall be in limitation of any other remedy, right, undertaking, or obligation of either party. Any breach by either party shall not be deemed material unless, within thirty (30) days (or fifteen (15) days for failure to pay monies owed) after the non-breaching party learns of such breach, the non-breaching party serves written notice thereof on the breaching party specifying the nature of the breach and the breaching party fails to cure such breach, if any, within thirty (30) days (15 days regarding payments) after receipt of such notice, or within a reasonable time thereafter if such breach is not curable within thirty (30) days. In entering into this agreement and providing services pursuant hereto, Artist has and shall have the status of an independent contractor and nothing herein contained shall contemplate or constitute Artist as Producers agent or employee. ARTIST UNDERSTANDS THAT ARTIST HAS THE RIGHT TO SEEK THE ADVICE OF INDEPENDENT COUNSEL CONCERNING ITS RIGHTS, THE PROVISIONS HEREOF, AND THE ADVISABILITY OF EXECUTING THIS LEGALLY BINDING AGREEMENT. FURTHER, ARTIST ACKNOWLEDGES THAT PRODUCER HAS GIVEN ARTIST THE OPPORTUNITY TO SEEK THE ADVICE OF INDEPENDENT COUNSEL AND ARTIST ACKNOWLEDGES THAT ARTIST IS EXECUTING THIS AGREEMENT VOLUNTARILY AFTER CONSULTATION WITH INDEPENDENT COUNSEL OR INTENTIONALLY DECIDING NOT TO SEEK ADVICE OF INDEPENDENT COUNSEL',
        '',
        'This agreement may be executed via facsimile and in two or more counterparts, each of which shall be deemed an original, but all of which shall constitute one instrument. In addition, a signed copy of this agreement transmitted by facsimile, by digital signature, or scanned into an image file and transmitted via email shall, for all purposes, be treated as if it was delivered containing an original manual signature of the party whose signature appears thereon and shall be binding upon such party as though an originally signed document had been delivered.',
        
        'Additional Note: Beats bought exclusively might have been leased out before being sold exclusively.'
      ].join('\n')
    };
  }

  return {
    title: 'Standard Beat License',
    content: [
      'This license grants the Licensee a non-exclusive right to use the purchased beat in a finished musical work under the applicable terms.',
      '',
      'The Licensee may use the beat for a finished work, may distribute that finished work, and may not resell or sublicense the beat as a standalone audio file.',
      '',
      'Please review the purchase details and contact De Beat Chef if you require any clarifications or modifications to this license.'
    ].join('\n')
  };
}

function buildLicenseAgreementText(customerName, customerEmail, orderDetails) {
  const timestamp = new Date().toISOString().split('T')[0];
  const orderLines = [];
  const uniqueLicenseTypes = new Set();

  if (Array.isArray(orderDetails)) {
    orderDetails.forEach(function(item, index) {
      const label = item.licenseType || item.license || 'Standard License';
      uniqueLicenseTypes.add(label);
      const template = findLicenseTemplate(label);

      orderLines.push('Item ' + (index + 1) + ': ' + (item.beat || 'Untitled Beat'));
      orderLines.push('License: ' + template.title);
      orderLines.push('License Fee: ' + (item.price != null ? '₦' + item.price : 'N/A'));
      orderLines.push('Included Files: ' + (Array.isArray(item.allowedFiles) ? item.allowedFiles.join(', ') : 'N/A'));
      orderLines.push('');
      orderLines.push(template.content.replace(/\[Beat Title\]/g, item.beat || 'Preview Track Only'));
      orderLines.push('');
      orderLines.push('---');
      orderLines.push('');
    });
  }

  const header = [
    'De Beat Chef License Agreement',
    'Generated Date: ' + timestamp,
    'Customer Name: ' + (customerName || 'N/A'),
    'Customer Email: ' + (customerEmail || 'N/A'),
    'License Document for Purchase items: ' + (Array.from(uniqueLicenseTypes).join(', ') || 'N/A'),
    '',
    'This document confirms the license terms for the beats purchased from De Beat Chef. The attached agreement reflects the license type purchased for each item and the applicable use rights.',
    '',
    'Purchased Item Details:',
    ''
  ];

  return header.concat(orderLines).join('\n');
}

function escapeHtml(text) {
  return text
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildLicenseAgreementAttachment(customerName, customerEmail, orderDetails) {
  const text = buildLicenseAgreementText(customerName, customerEmail, orderDetails);
  if (!text) {
    return null;
  }

  const filename = 'DeBeatChef-License-Agreement-' + (new Date().toISOString().split('T')[0]) + '.pdf';
  const htmlContent = '<html><head><meta charset="UTF-8"><style>body{font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:#111;margin:24px;}pre{white-space:pre-wrap;word-wrap:break-word;}h1{font-size:16pt;margin-bottom:12px;}p{margin:0 0 8px;} .section{margin-bottom:16px;}</style></head><body><h1>De Beat Chef License Agreement</h1><pre>' + escapeHtml(text) + '</pre></body></html>';
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(800).setHeight(1100);
  const pdfBlob = htmlOutput.getAs(MimeType.PDF).setName(filename);

  return pdfBlob;
}

function doPost(e) {
  try {
    Logger.log('doPost called with parameter object: %s', JSON.stringify(e && e.parameter ? e.parameter : {}));
    Logger.log('doPost raw body: %s', e && e.postData ? e.postData.contents : '(none)');
    const spreadsheetId = '10CzFZabv29Id_mg4kdHHDDT_QbAfZpxAPWsFdo-JohU';
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let params = e.parameter || {};
    if ((!params || Object.keys(params).length === 0) && e.postData && e.postData.contents) {
      try {
        params = Object.fromEntries(
          e.postData.contents.split('&').map(function(pair) {
            const parts = pair.split('=');
            return [decodeURIComponent(parts[0] || '').trim(), decodeURIComponent(parts.slice(1).join('=' ) || '').trim()];
          })
        );
      } catch (parseError) {
        Logger.log('doPost fallback parse error: %s', parseError && parseError.stack ? parseError.stack : parseError);
      }
    }
    Logger.log('doPost final params: %s', JSON.stringify(params));
    const type = (params.type || 'payment').toString().toLowerCase();
    const timestamp = params.timestamp || new Date().toISOString();

    if (type === 'payment') {
      try {
        const verifiedTransaction = verifyFlutterwavePayment(params.transactionId, params.amount, params.currency || 'NGN');
        params.status = 'success';
        params.paymentReference = verifiedTransaction.tx_ref || params.paymentReference || '';
        params.rawResponse = JSON.stringify(verifiedTransaction);
      } catch (verificationError) {
        return createCorsJsonOutput({ ok: false, message: verificationError.message });
      }
    }

  if (type === 'exclusive_offer') {
    const incomingFrontend = normalizeFrontendUrl(params.frontendUrl || params.frontend_url || '');
    if (incomingFrontend) {
      PropertiesService.getScriptProperties().setProperty('STORE_FRONTEND_URL', incomingFrontend);
    }
  }

  const propertyFrontendUrl = normalizeFrontendUrl(PropertiesService.getScriptProperties().getProperty('STORE_FRONTEND_URL'));

  let sheetName = 'Store Responses';
  let headers = [
    'timestamp',
    'type',
    'name',
    'email',
    'message',
    'beatTitle',
    'beatGenre',
    'beatBpm',
    'beatKey',
    'offerPrice',
    'offerMessage',
    'paymentReference',
    'amount',
    'currency',
    'status',
    'orderItems',
    'orderSummary',
    'rawResponse',
    'phone'
  ];

  if (type === 'contact') {
    sheetName = 'Contacts';
    headers = ['timestamp', 'type', 'name', 'email', 'message', 'phone'];
  } else if (type === 'exclusive_offer') {
    sheetName = 'Offers';
    headers = ['timestamp', 'type', 'name', 'email', 'customerEmail', 'adminEmail', 'scriptUrl', 'frontendUrl', 'itemId', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage', 'actionToken', 'status', 'actionTaken', 'actionTimestamp', 'payLinkToken', 'payLinkUrl'];
  } else if (type === 'order_request') {
    sheetName = 'Order Requests';
    headers = ['timestamp', 'type', 'name', 'email', 'amount', 'orderItems', 'orderSummary', 'downloadLinks', 'status'];
  } else {
    sheetName = 'Payments';
    headers = ['timestamp', 'type', 'name', 'email', 'paymentReference', 'amount', 'currency', 'status', 'orderItems', 'orderSummary', 'downloadLinks', 'rawResponse'];
  }

  if (type === 'account_create' || type === 'account_signin' || type === 'account_forgot_password') {
    const accountSheet = getAccountSheet(ss);
    const email = (params.email || '').toString().trim().toLowerCase();
    const password = (params.password || '').toString();
    const name = (params.name || '').toString().trim();

    if (type === 'account_create') {
      if (!name || !email || !password) {
        return createCorsJsonOutput({
          ok: false,
          message: 'Please provide your name, email, and password.'
        });
      }

      const existingAccount = findAccountByEmail(accountSheet, email);
      if (existingAccount) {
        return createCorsJsonOutput({
          ok: false,
          message: 'An account with this email already exists.'
        });
      }

      accountSheet.appendRow([timestamp, type, name, email, hashPassword(password), 'active']);
      return createCorsJsonOutput({
        ok: true,
        message: 'Account created successfully.',
        user: { name: name, email: email }
      });
    }

    if (type === 'account_forgot_password') {
      if (!email || !password) {
        return createCorsJsonOutput({
          ok: false,
          message: 'Please provide your email and new password.'
        });
      }

      const updated = updateAccountPassword(accountSheet, email, password);
      if (!updated) {
        return createCorsJsonOutput({
          ok: false,
          message: 'No account found with that email.'
        });
      }

      return createCorsJsonOutput({
        ok: true,
        message: 'Password reset successfully.'
      });
    }

    if (!email || !password) {
      return createCorsJsonOutput({
        ok: false,
        message: 'Please provide your email and password.'
      });
    }

    const account = findAccountByEmail(accountSheet, email);
    if (!account || account.password !== hashPassword(password)) {
      return createCorsJsonOutput({
        ok: false,
        message: 'Invalid email or password.'
      });
    }

    return createCorsJsonOutput({
      ok: true,
      message: 'Signed in successfully.',
      user: { name: account.name, email: account.email }
    });
  }

  const sheet = getOrCreateSheet(ss, sheetName);
  ensureHeaders(sheet, headers);

  const offerToken = type === 'exclusive_offer' ? generateSecureToken(48) : '';
  const requestFrontendUrl = normalizeFrontendUrl(params.frontendUrl || params.frontend_url || propertyFrontendUrl || '');
  if (requestFrontendUrl) {
    PropertiesService.getScriptProperties().setProperty('STORE_FRONTEND_URL', requestFrontendUrl);
  }

  const data = {
    timestamp: timestamp,
    type: type,
    name: params.name || '',
    email: params.email || params.offerEmail || '',
    customerEmail: params.customerEmail || params.email || params.offerEmail || '',
    adminEmail: params.adminEmail || params.admin_email || '',
    scriptUrl: params.scriptUrl || '',
    frontendUrl: requestFrontendUrl,
    itemId: params.itemId || '',
    beatTitle: params.beatTitle || '',
    beatGenre: params.beatGenre || '',
    beatBpm: params.beatBpm || '',
    beatKey: params.beatKey || '',
    offerPrice: params.offerPrice || '',
    offerMessage: params.offerMessage || '',
    actionToken: offerToken,
    status: type === 'exclusive_offer' ? 'pending' : params.status || '',
    actionTaken: '',
    actionTimestamp: '',
    payLinkToken: '',
    payLinkUrl: '',
    paymentReference: params.paymentReference || '',
    amount: params.amount || '',
    currency: params.currency || '',
    orderItems: params.orderItems || '',
    orderSummary: params.orderSummary || '',
    downloadLinks: params.downloadLinks || '',
    rawResponse: params.rawResponse || '',
    phone: params.phone || ''
  };

  const row = headers.map(function (header) {
    return data[header] || '';
  });

  sheet.appendRow(row);

  const customerEmail = (data.customerEmail || params.customerEmail || params.email || params.offerEmail || '').toString().trim();
  const adminEmail = (data.adminEmail || params.adminEmail || params.admin_email || DEFAULT_ADMIN_EMAIL).toString().trim();
  const scriptUrl = getWebAppUrl(data) || getWebAppUrl(params);
  const actionToken = data.actionToken || '';
  const customerName = (params.name || 'Customer').toString().trim() || 'Customer';
  const customerMessage = (params.message || '').toString();
  const notificationResults = [];

  if (type === 'exclusive_offer' && adminEmail) {
    const acceptUrl = `${scriptUrl}?type=offer_action&action=accept&token=${encodeURIComponent(actionToken)}`;
    const declineUrl = `${scriptUrl}?type=offer_action&action=decline&token=${encodeURIComponent(actionToken)}`;
    const sellerBody = [
      `<p>Hi,</p>`,
      `<p>You have received a new exclusive offer from <strong>${escapeHtml(customerEmail)}</strong> for <strong>${escapeHtml(params.beatTitle || 'Unknown Beat')}</strong>.</p>`,
      `<p><strong>Offer amount:</strong> ₦${escapeHtml(params.offerPrice || 'N/A')}</p>`,
      `<p><strong>Message:</strong></p>`,
      `<p>${escapeHtml(params.offerMessage || 'No message provided').replace(/\n/g, '<br>')}</p>`,
      '<p>Please choose one of the actions below:</p>',
      `<p><a href="${acceptUrl}" style="background:#16a34a;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;display:inline-block;margin-right:8px;">Accept offer</a><a href="${declineUrl}" style="background:#dc2626;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;display:inline-block;">Decline offer</a></p>`,
      `<p>Offer ID: ${escapeHtml(params.itemId || 'N/A')}</p>`
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: adminEmail,
      replyTo: customerEmail || '',
      subject: `New exclusive offer for ${params.beatTitle || 'your beat'}`,
      htmlBody: sellerBody
    }));

    if (customerEmail) {
      const customerOfferBody = [
        `<p>Hi ${escapeHtml(customerName)},</p>`,
        `<p>Your offer for <strong>${escapeHtml(params.beatTitle || 'the beat')}</strong> has been received.</p>`,
        `<p><strong>Offer amount:</strong> ₦${escapeHtml(params.offerPrice || 'N/A')}</p>`,
        '<p>The seller will review it and get back to you shortly.</p>',
        '<p>Best regards,<br>De Beat Chef</p>'
      ].join('');

      notificationResults.push(sendNotificationEmail({
        to: customerEmail,
        replyTo: adminEmail || 'jayomoluwa@gmail.com',
        subject: 'Your offer has been submitted',
        htmlBody: customerOfferBody
      }));
    }
  }

  if (type === 'contact' && customerEmail) {
    const customerBody = [
      `<p>Hi ${customerName},</p>`,
      '<p>Thanks for contacting De Beat Chef. We have received your message and will get back to you shortly.</p>',
      '<p><strong>Your message:</strong></p>',
      `<p>${customerMessage.replace(/\n/g, '<br>') || 'No message provided.'}</p>`,
      '<p>Best regards,<br>De Beat Chef</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: customerEmail,
      replyTo: adminEmail || customerEmail,
      subject: 'We received your message',
      htmlBody: customerBody
    }));
  }

  if (type === 'contact' && adminEmail) {
    const adminBody = [
      `<p>You have a new contact message from ${customerName}.</p>`,
      `<p><strong>Email:</strong> ${customerEmail || 'Not provided'}</p>`,
      '<p><strong>Message:</strong></p>',
      `<p>${customerMessage.replace(/\n/g, '<br>') || 'No message provided.'}</p>`,
      '<p>Please reply to the customer directly.</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: adminEmail,
      replyTo: customerEmail || '',
      subject: `New contact message from ${customerName}`,
      htmlBody: adminBody
    }));
  }

  if (type === 'order_request' && customerEmail) {
    let downloadHtml = '';
    let orderDetails = [];
    try {
      orderDetails = JSON.parse(params.orderItems || '[]');
    } catch (e) {
      orderDetails = [];
    }

    try {
      const parsedLinks = JSON.parse(params.downloadLinks || '[]');
      if (Array.isArray(parsedLinks) && parsedLinks.length) {
        downloadHtml = parsedLinks.map(function(item) {
          const links = item.downloadLinks || {};
          const allowed = Array.isArray(item.allowedFiles) && item.allowedFiles.length ? item.allowedFiles : Object.keys(links);
          const files = allowed.map(function(key) {
            if (!links[key]) return null;
            return `<li><strong>${String(key).toUpperCase()}:</strong> <a href="${links[key]}" target="_blank">Download</a></li>`;
          }).filter(Boolean).join('');
          return `<div style="margin-bottom:16px;"><p><strong>${item.beat}</strong> (${item.license})</p><ul style="margin:0;padding-left:18px;">${files}</ul></div>`;
        }).join('');
      }
    } catch (e) {
      downloadHtml = '<p>Order files will be provided after the order is confirmed.</p>';
    }

    const licenseAttachment = buildLicenseAgreementAttachment(customerName, customerEmail, orderDetails);

    const customerBody = [
      `<p>Hi ${customerName},</p>`,
      '<p>Thanks for your order request from De Beat Chef!</p>',
      `<p><strong>Requested total:</strong> ${params.amount ? '$' + params.amount : 'N/A'}</p>`,
      '<p>The seller will contact you to confirm the order and arrange completion.</p>',
      `${downloadHtml}`,
      '<p>If you have any questions, reply to this email and we will help you.</p>',
      '<p>Best regards,<br>De Beat Chef</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: customerEmail,
      replyTo: adminEmail || customerEmail,
      subject: 'We received your beat order request',
      htmlBody: customerBody,
      attachments: licenseAttachment ? [licenseAttachment] : []
    }));

    if (adminEmail) {
      const adminPaymentBody = [
        `<p>New beat order request from ${customerName} (${customerEmail}).</p>`,
        `<p><strong>Requested amount:</strong> $${params.amount || 'N/A'}</p>`,
        `<p><strong>Requested items:</strong></p>`,
        `<pre style="white-space:pre-wrap;">${params.orderItems || 'N/A'}</pre>`
      ].join('');
      notificationResults.push(sendNotificationEmail({
        to: adminEmail,
        replyTo: customerEmail || '',
        subject: `New beat order request from ${customerName}`,
        htmlBody: adminPaymentBody
      }));
    }
  }

  return createCorsJsonOutput({
    ok: true,
    message: 'Saved to sheet and sent notifications',
    type: type,
    sheetName: sheetName,
    notifications: notificationResults
  });
  } catch (error) {
    Logger.log('doPost error: %s', error && error.stack ? error.stack : error);
    return createCorsJsonOutput({
      ok: false,
      message: 'Server error in doPost: ' + (error && error.message ? error.message : 'unknown'),
      error: String(error)
    });
  }
}
