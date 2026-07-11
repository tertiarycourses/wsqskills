#!/usr/bin/env node
/**
 * wsqskills installer.
 *
 *   npx github:tertiarycourses/wsqskills            # install into this project's ./.claude
 *   npx github:tertiarycourses/wsqskills --user     # install into ~/.claude (all projects)
 *   npx github:tertiarycourses/wsqskills --force    # overwrite files that already exist
 *
 * Copies the WSQ courseware skills, slash commands, the courseware-qa agent, the push scripts
 * and the three courseware hooks into a .claude directory, and registers the hooks in settings.json.
 *
 * Existing files are NEVER overwritten without --force: a course repo may carry a
 * course-customised generator, and clobbering it would destroy work.
 */
const fs = require("fs");
const path = require("path");
const os = require("os");

const args = new Set(process.argv.slice(2));
const USER = args.has("--user") || args.has("-u");
const FORCE = args.has("--force") || args.has("-f");
const SRC = path.join(__dirname, "..");
const DEST = USER ? path.join(os.homedir(), ".claude") : path.join(process.cwd(), ".claude");

let copied = 0, skipped = 0;

function copyTree(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    if (entry.name === "__pycache__" || entry.name === ".DS_Store") continue;
    const s = path.join(src, entry.name);
    const d = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyTree(s, d);
    } else if (fs.existsSync(d) && !FORCE) {
      console.log(`  skip (exists): ${path.relative(DEST, d)}`);
      skipped++;
    } else {
      fs.copyFileSync(s, d);
      if (d.endsWith(".py")) fs.chmodSync(d, 0o755);
      copied++;
    }
  }
}

// The hooks: courseware pre-check, post-generation QA, and the TMS-push QA gate.
function hookCmd(file) {
  return USER
    ? `python3 "${path.join(os.homedir(), ".claude", "hooks", file)}"`
    : `python3 "$CLAUDE_PROJECT_DIR/.claude/hooks/${file}"`;
}

function registerHooks() {
  const settingsPath = path.join(DEST, "settings.json");
  let settings = {};
  if (fs.existsSync(settingsPath)) {
    try {
      settings = JSON.parse(fs.readFileSync(settingsPath, "utf8"));
    } catch {
      console.error(`  ! ${settingsPath} is not valid JSON — leaving it alone.`);
      console.error("    Add the hooks manually (see the README).");
      return;
    }
  }
  const hooks = (settings.hooks ||= {});
  const want = [
    ["PreToolUse", "Bash", "courseware-pre-hook.py"],
    ["PreToolUse", "Bash|Skill", "courseware-tms-push-hook.py"],
    ["PostToolUse", "Bash", "courseware-post-hook.py"],
  ];
  for (const [event, matcher, file] of want) {
    const list = (hooks[event] ||= []);
    if (JSON.stringify(list).includes(file)) continue; // already registered
    list.push({ matcher, hooks: [{ type: "command", command: hookCmd(file) }] });
    console.log(`  hook registered: ${event} → ${file}`);
  }
  fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2) + "\n");
}

console.log(`\nInstalling wsqskills into ${DEST}${FORCE ? " (force)" : ""}\n`);
for (const dir of ["skills", "commands", "agents", "hooks", "scripts"]) {
  const src = path.join(SRC, dir);
  if (fs.existsSync(src)) copyTree(src, path.join(DEST, dir));
}
registerHooks();

console.log(`\nDone — ${copied} file(s) installed, ${skipped} left untouched.`);
console.log(FORCE ? "" : "Re-run with --force to overwrite the skipped files.");
console.log("\nStart Claude Code in your course repo and run /courseware-qa to check it works.\n");
