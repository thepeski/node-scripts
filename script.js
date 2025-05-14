/* script.js */
// root script

// custom imports
import { Excel } from "./src/index.js";

// Excel class
const e = new Excel();
await e.open(["Book1"], "data/input");

e.fetch("A1");

await e.saveAll();
e.closeAll();