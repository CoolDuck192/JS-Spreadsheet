import "@testing-library/jest-dom/vitest";
import { configure } from "@testing-library/react";

// Role + accessible-name queries over the rendered grid are expensive (hundreds of
// gridcells), so the 1s default times out on slower machines even though the app
// updates within ~150ms. Give async queries more headroom.
configure({ asyncUtilTimeout: 4000 });
