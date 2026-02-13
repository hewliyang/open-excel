import { createRoot } from "react-dom/client";
import App from "./components/app";
import "./index.css";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(<App />);
});
