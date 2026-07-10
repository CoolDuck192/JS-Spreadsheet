import React from "react";
import ReactDOM from "react-dom/client";

// Manual review: corepack pnpm exec vite --host 0.0.0.0 --port 5173
async function mount() {
  const root = ReactDOM.createRoot(document.getElementById("root") as HTMLElement);
  if (window.location.pathname === "/datatable" || window.location.pathname === "/datatable/") {
    const [{ DataTableDemo }] = await Promise.all([
      import("./demo/DataTableDemo"),
      import("./styles/data-table.css")
    ]);
    root.render(<React.StrictMode><DataTableDemo /></React.StrictMode>);
    return;
  }

  const [{ default: App }] = await Promise.all([
    import("./App"),
    import("./App.css")
  ]);
  root.render(<React.StrictMode><App /></React.StrictMode>);
}

void mount();
