import { domain } from "process";
import { SigntaureResponse } from "../../components/HomePage/App";
import { SIGN_DEFAULT, SIGN_FORWARD, SIGN_INTERNAL, SIGN_REPLY } from "../../constants/SignatureConstants";

export const extractSignature = (signature: SigntaureResponse, composeType: string,inDomian:boolean|null): string => {
  console.log(inDomian,6);
  console.log(signature,7);
  console.log(composeType,8);
  if (composeType == "reply" && (inDomian==null||inDomian==false)) {
    const index = signature.data.findIndex((ele) => {
      if (ele.name === SIGN_REPLY) {
        console.log(ele.name,"12");
        return ele.html_content;
      } else {
        return null;
      }
    });
    if (index > -1) {
      console.log(index,"19");
      return signature.data[index].html_content;
    } else {
      console.log(index,"22");
      return signature.data[0].html_content;
    }
  } else if (composeType == "forward" && (inDomian==null||inDomian==false)) {
    const index = signature.data.findIndex((ele) => {
      if (ele.name === SIGN_FORWARD) {
        console.log(ele.name,"28");
        return ele.html_content;
      } else {
        return null;
      }
    });
    if (index > -1) {
      console.log(index,"35");
      return signature.data[index].html_content;
    } else {
      return signature.data[0].html_content;
    }
  } else if(composeType =="newMail" && (inDomian==null||inDomian==false)){
    const index = signature.data.findIndex((ele) => {
      if (ele.name === SIGN_DEFAULT) {
        console.log(ele.name,"43");
        return ele.html_content;
      } else {
        return null;
      }
    });
    if (index > -1) {
      console.log(index,"50");
      return signature.data[index].html_content;
    } else {
      console.log(index,"53");
      return null;
    }
  }
  else
  {
    if(inDomian!=null && inDomian){
    const index = signature.data.findIndex((ele) => {
      if (ele.name === SIGN_INTERNAL) {
        console.log(ele.name,"62");
        return ele.html_content;
      } else {
        return null;
      }
    });
    if (index > -1) {
      console.log(index,"69");
      return signature.data[index].html_content;
    } else {
      console.log(index);
      return null;
    }

  }
  else
  {
    return null;
  }
}
};

export async function setSignature(messageType: string) {
  //  var sign='<!DOCTYPE html><html><head><meta content="text/html; charset=utf-8" http-equiv=Content-Type> <body style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif;"><div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Amulveer Singh Walia, Sales</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #548dd4;"><b>MetaDesign Solutions Pvt Ltd.</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Your Trusted Software Services Partner</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>ISO 9001:2008 Certified</b></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>Mobile.</span><span style="text-align: right;">+91-981-145-4488</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>USA.</span><span style="text-align: right;">+1 (917) 728-1770</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>UK.</span><span style="text-align: right;">+44 (208) 089-9360</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>AUS.</span><span style="text-align: right;">+61 (731) 180-0630</span></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Skype: amulveer_singh_walia <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Email: <a href=mailto:amulveer@metadesignsolutions.co.uk target=_blank>amulveer@metadesignsolutions.co.uk</a> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Website: <a href=http://www.metadesignsolutions.com rel=noreferrer target=_blank>www.metadesignsolutions.com</a></div> <div style="border-style: hidden; padding-top: 10px; padding-bottom: 10px; white-space: nowrap;"><a href=https://metadesignsolutions.com/ target=_blank><img alt="MetaDesign Solutions" border=0 height=82 nosend=1 src=http://res.cloudinary.com/metadesign/image/upload/v1498639605/logo-metadesign_ttkzqs.png style="display: block; line-height: 0; font-size: 0;" width=200> </a> </div>';

  console.log(messageType,"87");
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

export async function composeSignature(data: SigntaureResponse ,inDomian :boolean|null) {
  console.log(inDomian,"100")
  await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const sign = extractSignature(data, asyncResult.value.composeType,inDomian);
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

export async function composeSignatureOnSend(data: SigntaureResponse ,inDomian :boolean|null,async:any) {
  console.log(inDomian,"100")
  await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const sign = extractSignature(data, asyncResult.value.composeType,inDomian);
      console.log(
        "getComposeTypeAsync succeeded with composeType: " +
          asyncResult.value.composeType +
          " and coercionType: " +
          asyncResult.value.coercionType
      );
      setSignatureOnSend(sign,async);
    } else {
      console.error(asyncResult.error);
    }
  });
}

export async function setSignatureOnSend(messageType: string,async :any) {
 var sign='<!DOCTYPE html><html><head><meta content="text/html; charset=utf-8" http-equiv=Content-Type> <body style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif;"><div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Amulveer Singh Walia, Sales</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #548dd4;"><b>MetaDesign Solutions Pvt Ltd.</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>Your Trusted Software Services Partner</b> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0; padding: 0; color: #953734;"><b>ISO 9001:2008 Certified</b></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>Mobile.</span><span style="text-align: right;">+91-981-145-4488</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>USA.</span><span style="text-align: right;">+1 (917) 728-1770</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>UK.</span><span style="text-align: right;">+44 (208) 089-9360</span> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; display: block; width: 70%; margin: 0 0 .3rem 0; padding: 0;" width=70%><span style="text-align: left; display: inline-block; width: 30%;" width=30%>AUS.</span><span style="text-align: right;">+61 (731) 180-0630</span></div> <div style="margin-bottom: 1.5rem;"><p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Skype: amulveer_singh_walia <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Email: <a href=mailto:amulveer@metadesignsolutions.co.uk target=_blank>amulveer@metadesignsolutions.co.uk</a> <p style="-webkit-text-size-adjust: 100%; font-size: 10.5pt; font-family: Tahoma,sans-serif; margin: 0 0 .3rem 0;">Website: <a href=http://www.metadesignsolutions.com rel=noreferrer target=_blank>www.metadesignsolutions.com</a></div> <div style="border-style: hidden; padding-top: 10px; padding-bottom: 10px; white-space: nowrap;"><a href=https://metadesignsolutions.com/ target=_blank><img alt="MetaDesign Solutions" border=0 height=82 nosend=1 src=http://res.cloudinary.com/metadesign/image/upload/v1498639605/logo-metadesign_ttkzqs.png style="display: block; line-height: 0; font-size: 0;" width=200> </a> </div>';
console.log(async);

  console.log(messageType,"87");
  
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

