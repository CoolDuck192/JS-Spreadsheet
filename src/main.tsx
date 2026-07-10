import React from "react";
import ReactDOM from "react-dom/client";
import { Spreadsheet } from "./react/Spreadsheet";
import "./App.css";

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <Spreadsheet />
  </React.StrictMode>
);
