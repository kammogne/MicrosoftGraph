/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The contents of the outbound email message that will be sent to the user
const emailContent = `<html><head> <meta http-equiv='Content-Type' content='text/html; charset=us-ascii'> <title></title> </head>
  <body style='font-family:calibri'> <p>Congratulations {{name}}!</p> <p>This is a message from Landry Kammogne</p> </body> </html>`;

/**
 * Returns the outbound email message content with the supplied name populated in the text.
 * @param {string} name The proper noun to use when addressing the email.
 * @param {string} sharingLink The sharing link to the file to embed in the email.
 * @return {string} the formatted email body
 */
function getEmailContent(name, sharingLink) {
  let bodyContent = emailContent.replace('{{name}}', name);
  bodyContent = bodyContent.replace('{{sharingLink}}', sharingLink);
  return bodyContent;
}

/**
 * Wraps the email's message content in the expected [soon-to-deserialized JSON] format.
 * @param {string} content The message body of the email message.
 * @param {string} recipient The email address to whom this message will be sent.
 * @return the message object to send over the wire
 */
function wrapEmail(content, recipient, file) {
  const attachments = [{
    '@odata.type': '#microsoft.graph.fileAttachment',
    ContentBytes: file,
    Name: 'mypic.jpg'
  }];
  const emailAsPayload = {
    Message: {
      Subject: 'Welcome to Microsoft Graph development with Node.js and the Microsoft Graph Connect sample',
      Body: {
        ContentType: 'HTML',
        Content: content
      },
      ToRecipients: [
        {
          EmailAddress: {
            Address: recipient
          }
        }
      ]
    },
    SaveToSentItems: true,
    Attachments: attachments
  };
  return emailAsPayload;
}

/**
 * Delegating method to wrap the formatted email message into a POST-able object
 * @param {string} name the name used to address the recipient
 * @param {string} recipient the email address to which the connect email will be sent
 */
function generateMailBody(name, recipient, sharingLink, file) {
  return wrapEmail(getEmailContent(name, sharingLink), recipient, file);
}

exports.generateMailBody = generateMailBody;
