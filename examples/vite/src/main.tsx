import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@simplishelf/opscards/style.css";
import "@simplishelf/sheet-ingest/style.css";
import App from "./App";
import "./example.css";

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("Could not find the example app root element.");
}

createRoot(rootElement).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
