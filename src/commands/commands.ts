/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { domain } from "process";
import { SigntaureResponse } from "../taskpane/components/HomePage/App";
import { getSigntaure } from "../taskpane/service/APIService/GetSignature";
import { composeSignature, composeSignatureOnSend, extractSignature, setSignature } from "../taskpane/service/SignatureServcie/SignatureServices";
import {deferred} from "promise-callbacks"

/* global global, Office, self, window */
var mailboxItem;
Office.onReady(() => {
  // If needed, Office.js is ready to be called
  mailboxItem = Office.context.mailbox.item;
});


// Office.initialize = function (reason) {
//     mailboxItem = Office.context.mailbox.item;
// }

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
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
      const promise = deferred();
      var isInDOmain = await getAllRecipients();
      console.log(isInDOmain,"64");
      getSigntaure().then(async (data) => {
        await composeSignatureOnSend(data, isInDOmain,asyncResult).then(function () {
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
  var flag

  let isInDOmain;
  let checkAllCC;
  let checkBCC

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
  // toRecipients(toRecipients)

      CheckDOmain.push(await getAllCC(toRecipients));//null//true//false
      if(CheckDOmain[0] || CheckDOmain[0]===null)
      {
        // flag=CheckDOmain?true:null;
        CheckDOmain.push(await getToCC(ccRecipients));//null//true//false
        if(CheckDOmain[1] || CheckDOmain[1]===null)
        { 
          // flag=CheckDOmain?true:null;
          CheckDOmain.push(await getBCC(bccRecipients));//null//true//false
          if(CheckDOmain[2] || CheckDOmain[2]===null)
          {
            // flag=(CheckDOmain || flag==true)?flag:null;
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
      // return CheckDOmain===null?(flag==true?true:null):CheckDOmain;
      // return (checkToCC==null && checkAllCC== null &&checkBCC==null)?null:()
  // var promiseToCC= new Promise((resolve)=>{
  //    toRecipients.getAsync(async function (asyncResult) {
  //     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
  //       console.log(asyncResult.error.message);
  //     }
  //     else {
  //       // Async call to get to-recipients of the item completed.
  //       // Display the email addresses of the to-recipients. 
  //       console.log('To-recipients of the item:');
  //       CheckDOmain = await displayAddresses(asyncResult);
        
  //     }
  //     if (!CheckDOmain) {
  //       console.log(CheckDOmain, "97");
  //       return resolve(CheckDOmain);
  //     }
  //   });
  // })

//   var promiseCC= new Promise((resolve)=>{
//   ccRecipients.getAsync(async (asyncResult) => {
//     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
//       console.log(asyncResult.error.message);
//     }
//     else {
//       // Async call to get cc-recipients of the item completed.
//       // Display the email addresses of the cc-recipients.
//       console.log('Cc-recipients of the item:');
//       console.log(Office.context.mailbox.userProfile.emailAddress);
//       CheckDOmain = await displayAddresses(asyncResult);
//     }
//     if (!CheckDOmain) {
//       console.log(CheckDOmain, "112");
//       return CheckDOmain;
//     }
//   });
// }


  // if (bccRecipients) {
  //   bccRecipients.getAsync(async function (asyncResult) {
  //     if (asyncResult.status == Office.AsyncResultStatus.Failed) {
  //       console.log(asyncResult.error.message);
  //     }
  //     else {
  //       // Async call to get bcc-recipients of the item completed.
  //       // Display the email addresses of the bcc-recipients.
  //       console.log('Bcc-recipients of the item:');
  //       CheckDOmain = await displayAddresses(asyncResult);
  //     }
  //     if (!CheckDOmain) {
  //       console.log(CheckDOmain, "128");
  //       return CheckDOmain;
  //     }
  //   }); // End getAsync for bcc-recipients.
   
  // }
  // console.log(CheckDOmain, "130");
  // return CheckDOmain;
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
  Office.context.mailbox.item.subject.setAsync(
    "Set by an event-based add-in!",
    {
      "asyncContext" : event
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() after all work is done.
      asyncResult.asyncContext.completed();
    });
  // var isInDOmain = getAllRecipients();
  // console.log(isInDOmain, "165");
  // getSigntaure().then(async (data) => {
  //   console.log(data);
  //   composeSignature(data, await isInDOmain)
  // });
  // let messageType = "";
  // console.log(event.source);

  // console.log(messageType);
}

function insertDefaultGist(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
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

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.insertDefaultGist = insertDefaultGist;



