/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { getSendSignature, getSigntaure } from "../taskpane/service/APIService/GetSignature";
import { composeSignature, composeSignatureOnSend, extractSignature, setSignature } from "../taskpane/service/SignatureServcie/SignatureServices";


/* global global, Office, self, window */
var mailboxItem;
Office.onReady(() => {
  // If needed, Office.js is ready to be called
  mailboxItem = Office.context.mailbox.item;
});

// const dataContextStore=useContext(DataContext);
// Office.initialize = function (reason) {
//     mailboxItem = Office.context.mailbox.item;
// }

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

 

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function onMessageComposeHandler(event) {
  console.log("message composing")
  console.log(event)
  setSubject(event);
}
function onAppointmentComposeHandler(event) {
  console.log(event);
  // setSubject(event); 
}
function validateSubjectAndCC(event) {
  shouldChangeSubjectOnSend(event);
}


function shouldChangeSubjectOnSend(event) {
  mailboxItem.subject.getAsync(
    { asyncContext: event },
    async function (asyncResult) {
      console.log(asyncResult, "on send");
      var isInDOmain = await getAllRecipients();
      console.log(isInDOmain,"64");
      getSendSignature().then(async (data) => {
      // localStorage.setItem("AbbreViatedSignature", data.use_abbreviated_REPLY_signature.toString());
        console.log(data.data,"68");
        await composeSignatureOnSend(data.data, isInDOmain,asyncResult).then(function () {
        });
      });
    }
  )
}


export function getAllCC(toRecipients){
  return new Promise((resolve,reject)=>{
    toRecipients.getAsync(async function (asyncResult) {
     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
       console.log(asyncResult.error.message);
       reject(asyncResult.error.message);
     }
     else {
       
       // Async call to get to-recipients of the item completed.
       // Display the email addresses of the to-recipients. 
       console.log('To-recipients of the item:');
        const CheckDOmain = await displayAddresses(asyncResult);
       resolve(CheckDOmain)   
     }
   });
 })
}


export function getToCC(ccRecipients){
  return new Promise((resolve,reject)=>{
    ccRecipients.getAsync(async (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        reject(asyncResult.error.message);
      }
      else {
        // Async call to get cc-recipients of the item completed.
        // Display the email addresses of the cc-recipients.
        console.log('Cc-recipients of the item:');
        console.log(Office.context.mailbox.userProfile.emailAddress);
       const CheckDOmain = await displayAddresses(asyncResult);
        resolve(CheckDOmain);
      }
    });
  });
}


export function getBCC(bccRecipients){
  return new Promise((resolve,reject)=>{
    bccRecipients.getAsync(async (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        reject(asyncResult.error.message);
      }
      else {
        // Async call to get cc-recipients of the item completed.
        // Display the email addresses of the cc-recipients.
        console.log('Cc-recipients of the item:');
        console.log(Office.context.mailbox.userProfile.emailAddress);
       const CheckDOmain = await displayAddresses(asyncResult);
        resolve(CheckDOmain);
      }
    });
  });
}


export async function getAllRecipients() {

  var CheckDOmain=[];
  let isInDOmain;

  var toRecipients, ccRecipients, bccRecipients;
  if (mailboxItem.itemType == Office.MailboxEnums.ItemType.Appointment) {
    toRecipients = mailboxItem.requiredAttendees;
    ccRecipients = mailboxItem.optionalAttendees;
  }
  else {
    toRecipients = mailboxItem.to;
    ccRecipients = mailboxItem.cc;
    bccRecipients = mailboxItem.bcc;
  }

      CheckDOmain.push(await getAllCC(toRecipients));//null//true//false
      if(CheckDOmain[0] || CheckDOmain[0]===null)
      {
        CheckDOmain.push(await getToCC(ccRecipients));//null//true//false
        if(CheckDOmain[1] || CheckDOmain[1]===null)
        { 
          CheckDOmain.push(await getBCC(bccRecipients));//null//true//false
          if(CheckDOmain[2] || CheckDOmain[2]===null)
          {

          }
        }
      }
      isInDOmain = CheckDOmain.find((ele)=>{
        return ele===false;
      });
      if(isInDOmain===undefined)
      {
        isInDOmain=CheckDOmain.find((ele)=>{
          return ele===true;
        });
      }
      return isInDOmain===undefined?null:isInDOmain;
}


function displayAddresses(asyncResult) {
  var IsSameDomain = null;
  var CurrentUserAddress = Office.context.mailbox.userProfile.emailAddress;
  // var CurrentUserDomain = CurrentUserAddress.split('@')[1];
  var CurrentUserDomain="metadesignsolutions.co.uk"
  console.log(asyncResult.value.length, "146");
  for (var i = 0; i < asyncResult.value.length; i++) {
    console.log(asyncResult.value[i].emailAddress, "148")
    var RecpDomain = asyncResult.value[i].emailAddress.split('@')[1];
    console.log(RecpDomain,"230");
    if (CurrentUserDomain == RecpDomain) {
      IsSameDomain = true;
    }
    else {
      IsSameDomain = false;
      break;
    }
  }
  console.log(IsSameDomain, "158");
  return IsSameDomain;
}

async function setSubject(event) {
  var isInDOmain = getAllRecipients();
  console.log(isInDOmain, "165");
  getSigntaure().then(async (data) => {
    console.log(data);
    composeSignature(data.data, await isInDOmain)
  });
  // Office.context.mailbox.item.subject.setAsync(
  //   "Set by an event-based add-in!",
  //   {
  //     "asyncContext": event
  //   },
  //   function (asyncResult) {
  //     // Handle success or error.
  //     if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
  //       console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
  //     }

  //     // Call event.completed() after all work is done.
  //     asyncResult.asyncContext.completed();
  //   });
}

function insertDefaultGist(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox.item.body.setSignatureAsync(
    "Ravi Kumar\n MDS\n Front-end Developer",
    {
      coercionType: Office.CoercionType.Html,
    },

    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
    }
  );
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
Office.actions.associate("validateSubjectAndCC", validateSubjectAndCC);

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
g.insertDefaultGist = insertDefaultGist;



