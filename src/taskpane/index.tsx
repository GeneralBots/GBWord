import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global Word, document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "General Bots";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

function handler() {}

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;

  // Add the event handler.
  Word.run(async (context) => {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    return context.sync();
  });
  render(App);
};

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
