// Entry của FunctionFile: đăng ký mọi ribbon command (ExecuteFunction) từ các nhóm trong src/features/.
// Thêm nhóm mới: tạo src/features/<group>/index.js rồi thêm vào GROUPS.
// Key phải khớp <FunctionName> trong manifest/manifest.template.xml và không trùng giữa các nhóm.

import { makeCommand } from "./runCommand.js";
import dgnl from "../features/dgnl/index.js";

const GROUPS = { dgnl };

Office.onReady(() => {
  const seen = new Set();
  for (const [group, commands] of Object.entries(GROUPS)) {
    for (const [name, { label, run }] of Object.entries(commands)) {
      if (seen.has(name)) throw new Error(`Duplicate command "${name}" in group "${group}"`);
      seen.add(name);
      Office.actions.associate(name, makeCommand(label, run));
    }
  }
});
