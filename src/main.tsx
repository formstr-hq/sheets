import { createRoot } from "react-dom/client";
import "./styles.css";
import App from "./App";
import { bootstrapLocalRelay } from "./lib/localRelay";

bootstrapLocalRelay();

createRoot(document.getElementById("root")!).render(<App />);
