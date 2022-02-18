import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { Header } from "../Header/Header";
import {Progress} from "../ProgressBar/Progress";
import { getSigntaure } from "../../service/APIService/GetSignature";
import { composeSignature } from "../../service/SignatureServcie/SignatureServices";
import { getAllRecipients } from "../../../commands/commands";
import "./App.scss"

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}


export interface SigntaureResponse{
  data: resp[];
  
}

export interface resp{
name :string,
html_content:string,
filename :string
}

  const App =(props:any)=> {

  // const [ signature, setSigntaure]=React.useState(''); 
  const { title, isOfficeInitialized } = props;

    const clicktoDelete = async () => {
      Office.context.mailbox.item.body.setSignatureAsync(
        "",
        {
          coercionType: Office.CoercionType.Text,
        },
        function (asyncResult) {
          console.log("set_signature - " + JSON.stringify(asyncResult));
        }
      );
    };

    const clicktoAdd = async () => {
      
      getSigntaure().then((response:SigntaureResponse)=>{
        console.log(response);
        var isInDOmain= getAllRecipients();
        console.log(isInDOmain);
         getSigntaure().then((data) => {
          console.log(data);
          //  composeSignature(data,isInDOmain)
       });
      });
    }

    const clicktogetUserName = async () => {
      Office.context.mailbox.item.body.getTypeAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
        } else {
            console.log("Body type: " + asyncResult.value,"63");
        }
      });
    };

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo={require("./../../../../assets/big-logo.png")} title={props.title} message="Welcome" /> */}
          <DefaultButton className="ms-welcome__action ButtonClass" iconProps={{ iconName: "ChevronRight" }} onClick={clicktogetUserName}>
            Create / Edit Signature
          </DefaultButton>
          {/* <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={clicktoDelete}>
            delete
          </DefaultButton>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={clicktogetUserName}>
            Get
          </DefaultButton> */}
      </div>
    );
  
}

export {App}
