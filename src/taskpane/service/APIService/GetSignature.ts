import { SigntaureResponse } from "../../components/HomePage/App";

export const getSigntaure=():Promise<SigntaureResponse>=> {
  var myHeaders = new Headers();
  myHeaders.append("Accept", "application/json");

  // var userName=Office.context.mailbox.userProfile.emailAddress;
  var userName="ravi@metadesignsolutions.co.uk";
  // return fetch("https://manage.dynasend.net/s/outlook-bundle?email={amulveer@metadesignsolutions.co.uk}", { method: 'get',headers: myHeaders,redirect: 'follow'})
  return fetch("https://manage.dynasend.net/s/outlook-bundle?email="+userName+"", { method: 'get',headers: myHeaders,redirect: 'follow'})
      .then(response=>response.json());
}