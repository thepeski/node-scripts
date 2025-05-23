/* script.js */
// root script

// custom imports
import { Excel } from "./src/index.js";

// Excel class
const e = new Excel();
await e.open(["Book1"], "data/input");

let copied = e.fetchRange("A1:B10", "formula");
// console.log(copied);
e.setRange("D1", copied, "formula");
// e.setRange("F8", copied, "formula");

await e.saveAll();
e.closeAll();