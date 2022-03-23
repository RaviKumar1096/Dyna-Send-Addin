import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { Header } from "../Header/Header";
import {Progress} from "../ProgressBar/Progress";
import { getSigntaure } from "../../service/APIService/GetSignature";
import { composeSignature } from "../../service/SignatureServcie/SignatureServices";
import { getAllRecipients } from "../../../commands/commands";
import "./App.scss"
import { DataContext, SigntaureResponseState } from "../../Store/Store";
import { useRecoilState } from "recoil";



/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}


export interface SigntaureResponse{
  data: resp[];
  use_abbreviated_REPLY_signature?:boolean;
  flag?:boolean
}

export interface resp{
name :string,
html_content?:string,
filename :string
text_content? :string

}

  const App = (props: any) => {
    const { title, isOfficeInitialized } = props;
    const [signatureResponse, setSignatureResponse] = React.useState<SigntaureResponse>(null);
    const [checkedBox, setCheckBox] = React.useState<boolean>(false);
    const [signButton, setSignButton] = React.useState<boolean>(false);
    const [disablebtn ,setDisablebtn]=React.useState<boolean>(false);

    // const [signatuerResponseData, setSignatuerResponseData] = useRecoilState(SigntaureResponseState);

    React.useEffect(()=>{
        console.log(signatureResponse);
        // dataContextStore.setSignatureResponse(data);
        // setSignatureResponse(data);
        // DataContext.Provider(value={signatureResponse});

    },[signatureResponse]);



    React.useEffect(()=>{
      getSigntaure().then(async (data:SigntaureResponse) => {
        setSignatureResponse(data);
        console.log(signatureResponse);
        setCheckBox(data.use_abbreviated_REPLY_signature);
        setDisablebtn(data.use_abbreviated_REPLY_signature);
        setSignButton(data.flag);
        localStorage.setItem("AbbreViatedSignature", data.use_abbreviated_REPLY_signature.toString());
        // dataContextStore.setSignatureResponse(data);
        // setSignatureResponse(data);
        // DataContext.Provider(value={signatureResponse});
      });

    },[])

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
      getSigntaure().then((response: SigntaureResponse) => {
        console.log(response);
        var isInDOmain = getAllRecipients();
        console.log(isInDOmain);
        getSigntaure().then((data) => {
          console.log(data);
          //  composeSignature(data,isInDOmain)
        });
      });
    };

    const clicktogetUserName = async () => {
    window.open("https://manage.dynasend.net/mysignature");
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

    const handleChange=(e:React.ChangeEvent<HTMLInputElement>)=>{

      console.log(e.target.checked);
      setCheckBox(e.target.checked);
      localStorage.setItem("AbbreViatedSignature", e.target.checked.toString());
      var item_value = localStorage.getItem("AbbreViatedSignature");
      console.log(item_value);
    }


    return (
      <div className="signatureConatiner">
        {signatureResponse===null?<></>:
        <>      
      <div className="ButtonConatiner">
        <DefaultButton
          className="ms-welcome__action ButtonClass"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={clicktogetUserName}
          disabled={!signButton}
        >
          Create / Edit Signature
        </DefaultButton>
        </div>

        <div className="validationContainer">
          <input type="checkbox" id="AbbSign" onChange={(e)=>handleChange(e)} checked={checkedBox} disabled={!disablebtn}/>
          use abbreviated REPLY signature
        </div>
          </>
        }
      </div>
    );
  };

export {App}
