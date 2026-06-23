import { createRoot } from "react-dom/client";
import { Switch, Route } from "wouter";
import App from "./App";
import AdminPage from "./pages/admin";
import NotFound from "./pages/not-found";
import "./index.css";

declare const Office: typeof import("@microsoft/office-js");

function Root() {
  return (
    <Switch>
      <Route path="/admin" component={AdminPage} />
      <Route path="/" component={App} />
      <Route component={NotFound} />
    </Switch>
  );
}

function mountApp() {
  const root = document.getElementById("root")!;
  createRoot(root).render(<Root />);
}

// If Office.js is available (running inside Excel), wait for it to initialize.
// Otherwise mount immediately (standalone preview mode).
if (typeof Office !== "undefined" && Office.initialize !== undefined) {
  Office.onReady(() => mountApp());
} else {
  mountApp();
}
