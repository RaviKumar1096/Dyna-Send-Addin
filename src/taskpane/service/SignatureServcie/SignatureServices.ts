import { resp, SigntaureResponse } from "../../components/HomePage/App";
import { ABBREVIATED_SIGN, COMPOSE_FORWARD, COMPOSE_NEWMAIL, COMPOSE_REPLY, ErrorMsg, SIGN_DEFAULT, SIGN_FORWARD, SIGN_INTERNAL, SIGN_REPLY } from "../../constants/SignatureConstants";

export const extractSignature = (signature: resp[], composeType: string, inDomian: boolean | null,onsend:boolean): string => {
  console.log(inDomian, 6);
  console.log(signature, 7);
  console.log(composeType, 8);
  const isAbbreviated = localStorage.getItem(ABBREVIATED_SIGN);
  console.log(isAbbreviated, "10");

  const internalSignatue = signature.filter((x) => x.name === SIGN_INTERNAL);
  const defaultSignature = signature.filter((x) => x.name === SIGN_DEFAULT);
  const replySignature = signature.filter((x) => x.name === SIGN_REPLY);

  if (defaultSignature.length > 0 && internalSignatue.length < 0 && replySignature.length < 0) {
    return defaultSignature[0].html_content;
  } else if (defaultSignature.length > 0 && replySignature.length > 0 && internalSignatue.length < 0) {
    if (composeType === COMPOSE_NEWMAIL) {
      return defaultSignature[0].html_content;
    } else {
      return replySignature[0].html_content;
    }
  } else if (defaultSignature.length > 0 && internalSignatue.length > 0 && replySignature.length > 0) {
    if(onsend==false)
    {
      return defaultSignature[0].html_content;
    }
    else if (inDomian != null && inDomian == true ) {
      return internalSignatue[0].html_content;
    } else if (composeType == COMPOSE_NEWMAIL && inDomian != null && inDomian == false) {
      return defaultSignature[0].html_content;
    } else if (
      (composeType === COMPOSE_REPLY || composeType == COMPOSE_FORWARD) &&
      inDomian != null &&
      inDomian == false &&
      isAbbreviated != "false"
    ) {
      return replySignature[0].html_content;
    } else {
      return defaultSignature[0].html_content;
    }
  } else if (defaultSignature.length > 0 && internalSignatue.length > 0 && replySignature.length < 0) {
    if (inDomian != null && inDomian == true) {
      return internalSignatue[0].html_content;
    } else {
      return defaultSignature[0].html_content;
    }
  } else {
    return ErrorMsg;
  }
};

export async function setSignature(messageType: string) {
  //  var sign='<!DOCTYPE html><html><head><meta content="text/html; charset=utf-8" http-equiv=Content-Type> <body style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif;"><div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Amulveer Singh Walia, Sales</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #548dd4;"><b>MetaDesign Solutions Pvt Ltd.</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Your Trusted Software Services Partner</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>ISO 9001:2008 Certified</b></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>Mobile.</span><span style="text-align: right;">+91-981-145-4488</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>USA.</span><span style="text-align: right;">+1 (917) 728-1770</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>UK.</span><span style="text-align: right;">+44 (208) 089-9360</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>AUS.</span><span style="text-align: right;">+61 (731) 180-0630</span></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Skype: amulveer_singh_walia <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Email: <a href=mailto:amulveer@metadesignsolutions.co.uk target=_blank>amulveer@metadesignsolutions.co.uk</a> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Website: <a href=http://www.metadesignsolutions.com rel=noreferrer target=_blank>www.metadesignsolutions.com</a></div> <div style="border-style: hidden; padding-top: 10px; padding-bottom: 10px; white-space: nowrap;"><a href=https://metadesignsolutions.com/ target=_blank><img alt="MetaDesign Solutions" border=0 height=82 nosend=1 src=http://res.cloudinary.com/metadesign/image/upload/v1498639605/logo-metadesign_ttkzqs.png style="display: block; line-height: 0; font-size: 0;" width=200> </a> </div>';

  console.log(messageType, "87");
  await Office.context.mailbox.item.body.setSignatureAsync(
    messageType,
    {
      coercionType: Office.CoercionType.Html,
    },
    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
    }
  );
}

export async function composeSignature(data: resp[], inDomian: boolean | null) {
  console.log(inDomian, "100");
  await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const sign = extractSignature(data, asyncResult.value.composeType, inDomian,false);
      console.log(
        "getComposeTypeAsync succeeded with composeType: " +
          asyncResult.value.composeType +
          " and coercionType: " +
          asyncResult.value.coercionType
      );
      setSignature(sign);
    } else {
      console.error(asyncResult.error);
    }
  });
}

export async function composeSignatureOnSend(data: resp[], inDomian: boolean | null, async: any) {
  console.log(data, inDomian, "100");
  var onsend=true;
  await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const sign = extractSignature(data, asyncResult.value.composeType, inDomian,onsend);
      console.log(
        "getComposeTypeAsync succeeded with composeType: " +
          asyncResult.value.composeType +
          " and coercionType: " +
          asyncResult.value.coercionType
      );
      setSignatureOnSend(sign, async);
    } else {
      console.error(asyncResult.error);
    }
  });
}

export async function setSignatureOnSend(messageType: string, async: any) {
  var sign =
    '<!DOCTYPE html><html><head><meta content="text/html; charset=utf-8" http-equiv=Content-Type> <body style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif;"><div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Amulveer Singh Walia, Sales</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #548dd4;"><b>MetaDesign Solutions Pvt Ltd.</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Your Trusted Software Services Partner</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>ISO 9001:2008 Certified</b></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>Mobile.</span><span style="text-align: right;">+91-981-145-4488</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>USA.</span><span style="text-align: right;">+1 (917) 728-1770</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>UK.</span><span style="text-align: right;">+44 (208) 089-9360</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>AUS.</span><span style="text-align: right;">+61 (731) 180-0630</span></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Skype: amulveer_singh_walia <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Email: <a href=mailto:amulveer@metadesignsolutions.co.uk target=_blank>amulveer@metadesignsolutions.co.uk</a> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Website: <a href=http://www.metadesignsolutions.com rel=noreferrer target=_blank>www.metadesignsolutions.com</a></div> <div style="border-style: hidden; padding-top: 10px; padding-bottom: 10px; white-space: nowrap;"><a href=https://metadesignsolutions.com/ target=_blank><img alt="MetaDesign Solutions" border=0 height=82 nosend=1 src=http://res.cloudinary.com/metadesign/image/upload/v1498639605/logo-metadesign_ttkzqs.png style="display: block; line-height: 0; font-size: 0;" width=200> </a> </div>';
  console.log(async);

  console.log(messageType, "87");

  await Office.context.mailbox.item.body.setSignatureAsync(
    messageType,
    {
      coercionType: Office.CoercionType.Html,
    },
    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
      async.asyncContext.completed({ allowEvent: true });
    }
  );
}
