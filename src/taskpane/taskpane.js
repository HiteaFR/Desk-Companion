/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

var config = {};

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
  config.zohodeskemail = Office.context.roamingSettings.get('zohodesk-email');
  if (config.zohodeskemail) {
    document.getElementById('zoho-email').value = config.zohodeskemail;
  }
});


export async function run() {

  config.zohodeskemail = document.getElementById('zoho-email').value;
  Office.context.roamingSettings.set('zohodesk-email', config.zohodeskemail);
  Office.context.roamingSettings.saveAsync();

  var item = Office.context.mailbox.item;

  Office.context.mailbox.item.body.getAsync(
    "html", {
      asyncContext: 'To Zoho Desk'
    },
    function callback(result) {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [config.zohodeskemail],
        subject: item.subject,
        htmlBody: '#original_sender{' + item.sender.emailAddress + '} <br/><hr><br/>' + result.value
      });
    });

}