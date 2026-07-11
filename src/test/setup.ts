import "@testing-library/jest-dom/vitest";
import { configure } from "@testing-library/react";

// Role + accessible-name queries over the rendered grid are expensive (hundreds of
// gridcells), so the 1s default times out on slower machines even though the app
// updates within ~150ms. Give async queries more headroom.
configure({ asyncUtilTimeout: 4000 });

if (typeof HTMLDialogElement !== "undefined" && !("showModal" in HTMLDialogElement.prototype)) {
  Object.defineProperty(HTMLDialogElement.prototype, "showModal", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.setAttribute("open", "");
    }
  });
}

if (typeof HTMLDialogElement !== "undefined" && !("close" in HTMLDialogElement.prototype)) {
  Object.defineProperty(HTMLDialogElement.prototype, "close", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.removeAttribute("open");
      this.dispatchEvent(new Event("close"));
    }
  });
}

if (!("showPopover" in HTMLElement.prototype)) {
  Object.defineProperty(HTMLElement.prototype, "showPopover", {
    configurable: true,
    value(this: HTMLElement) {
      this.setAttribute("data-popover-open", "true");
      this.style.display = "grid";
    }
  });
}

if (!("hidePopover" in HTMLElement.prototype)) {
  Object.defineProperty(HTMLElement.prototype, "hidePopover", {
    configurable: true,
    value(this: HTMLElement) {
      this.removeAttribute("data-popover-open");
      this.style.display = "none";
    }
  });
}
