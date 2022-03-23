import React, { createContext } from "react";
import { atom } from "recoil";
import { SigntaureResponse } from "../components/HomePage/App";

interface IDataContext {
    signatureResponse:SigntaureResponse,
    setSignatureResponse:React.Dispatch<React.SetStateAction<SigntaureResponse>>
}


export const SigntaureResponseState = atom({
    key: 'SigntaureResponse', // unique ID (with respect to other atoms/selectors)
    default: {}, // default value (aka initial value)
  });

export const DataContext = createContext({} as IDataContext)