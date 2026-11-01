import { createRoot } from "react-dom/client";
import { locale } from "@fortune-sheet/core";
import "@fortune-sheet/react/dist/index.css";
import "./styles.css";
import App from "./App";
import { bootstrapLocalRelay } from "./lib/localRelay";
import { DEFAULT_FONT_FAMILIES } from "./lib/workbook";

const en = locale({ lang: "en" } as never);
en.fontarray.splice(0, en.fontarray.length, ...DEFAULT_FONT_FAMILIES);
en.fontjson = { ...en.fontjson, inter: 0, "times new roman": 4 } as typeof en.fontjson;

bootstrapLocalRelay();

createRoot(document.getElementById("root")!).render(<App />);
