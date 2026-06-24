import { createRoot } from "react-dom/client";
import { Switch, Route } from "wouter";
import { AppConfigProvider } from "./context/AppConfigContext";
import App from "./App";
import AdminPage from "./pages/admin";
import EulaPage from "./pages/eula";
import PrivacyPage from "./pages/privacy";
import SupportPage from "./pages/support";
import NotFound from "./pages/not-found";
import "./index.css";

declare const Office: typeof import("@microsoft/office-js");

function Root() {
  return (
    <AppConfigProvider>
      <Switch>
        <Route path="/admin" component={AdminPage} />
        <Route path="/eula" component={EulaPage} />
        <Route path="/privacy" component={PrivacyPage} />
        <Route path="/support" component={SupportPage} />
        <Route path="/" component={App} />
        <Route component={NotFound} />
      </Switch>
    </AppConfigProvider>
  );
}

function mountApp() {
  const root = document.getElementById("root")!;
  createRoot(root).render(<Root />);
}

if (typeof Office !== "undefined" && Office.initialize !== undefined) {
  Office.onReady(() => mountApp());
} else {
  mountApp();
}
