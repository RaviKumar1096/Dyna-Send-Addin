import {App, SigntaureResponse} from "./components/HomePage/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { appendFunction, ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import {DataContext} from "./Store/Store"
import RecoilNexus from "recoil-nexus";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (App) => {

  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <App title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,

    document.getElementById("container")
  );
};

/* Render application after Office initializes */
// function Office.onReady(render(App) (info: {
//   host: Office.HostType;
//   platform: Office.PlatformType;
// }) => any): Promise<{
//   host: Office.HostType;
//   platform: Office.PlatformType;
// }>

// Office.onReady(render(App)=>{
//   isOfficeInitialized = true;
// });
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

if ((module as any).hot) {
  (module as any).hot.accept("./components/HomePage/App", () => {
    const NextApp = require("./components/HomePage/App").default;
    render(NextApp);
  });
}
