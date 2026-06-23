import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

declare const Office: typeof import("@microsoft/office-js");

function mountApp() {
  const root = document.getElementById("root")!;
  createRoot(root).render(<App />);
}

// If Office.js is available (running inside Excel), wait for it to initialize.
// Otherwise mount immediately (standalone preview mode).
if (typeof Office !== "undefined" && Office.initialize !== undefined) {
  Office.onReady(() => mountApp());
} else {
  mountApp();
}
