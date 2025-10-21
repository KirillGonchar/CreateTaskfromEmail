/*
AUTHOR DETAILS:
* Author: Kyrylo Honchar
* Date: 21-Oct-2025
*/


Office.onReady(() => {
  document.getElementById('createTaskBtn').addEventListener('click', createTaskFromEmail);
});

async function createTaskFromEmail() {
  const status = (msg) => document.getElementById('status').innerText = msg;
  status('Reading email...');

  try {
    const item = Office.context.mailbox.item;

    // Get subject
    const subject = item.subject || '';

    // Get sender
    const from = item.from || {};
    const fromEmail = (from.emailAddress || from.address || '') ;
    const fromName = from.displayName || fromNameFromAddress(fromEmail);

    // Get HTML body
    const bodyHtml = await new Promise((resolve, reject) => {
      item.body.getAsync("html", function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve(asyncResult.value);
        } else {
          reject(asyncResult.error);
        }
      });
    });

    // Get internet message id (if available)
    const messageId = item.internetMessageId || item.id || '';

    // Build payload
    const payload = {
      subject: subject,
      bodyHtml: bodyHtml,
      fromEmail: fromEmail,
      fromName: fromName,
      messageId: messageId,
      receivedAt: (new Date(item.dateTimeCreated || Date.now())).toISOString()
    };

    status('Sending to Flow...');

    // Call the flow endpoint
    const flowUrl = 'https://6e21e2ddf83ae861bdff96f8c5342f.a0.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/783c6d76a1444bdd9584e1b40d8aa9e8/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ak-_NA4npMUP38XlZqm_hWsn87VOW-_Uv3lGgeP5R_8'; 

    const resp = await fetch(flowUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (!resp.ok) {
      const txt = await resp.text();
      status('Flow returned error: ' + resp.status + ' - ' + txt);
      return;
    }
      status('Sent successfully: ' + resp.status + ' - ' + txt);
  } catch (err) {
    status('Error: ' + (err.message || JSON.stringify(err)));
  }
}

function fromNameFromAddress(email) {
  if (!email) return '';
  const name = email.split('@')[0];
  const splitted = name.replace(/\./g, ' ').split('.').filter(Boolean);
  if (splitted.length == 1) return capitalizeFirstLetter(splitted[0]);
	const capitalized = splitted.map(capitalizeFirstLetter);
	const [firstName, ...rest] = capitalized;
	const lastName = rest.join(' ');
	return `${firstName} ${lastName}`;
}

function capitalizeFirstLetter(word)
{
	if (typeof word !== 'string') return '';
	if (!word.length) return '';
	return word.charAt(0).toUpperCase() + word.slice(1);
}
