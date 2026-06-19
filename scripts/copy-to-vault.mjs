import { existsSync, cpSync, readFileSync } from "fs";
import { join, dirname } from "path";
import { homedir } from "os";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(__dirname, "..");
const pluginDir = join(projectRoot, "md-to-docx");

const manifest = JSON.parse(
  readFileSync(join(pluginDir, "manifest.json"), "utf-8")
);
const pluginId = manifest.id;

// Try common vault locations
const possiblePaths = [
  join(homedir(), "Documents", "Obsidian", ".obsidian", "plugins", pluginId),
  join(homedir(), "Obsidian Vault", ".obsidian", "plugins", pluginId),
];

let target = possiblePaths.find((p) => existsSync(join(p, "..")));
if (!target) {
  const vaultPath = process.env.OBSIDIAN_VAULT;
  if (vaultPath) {
    target = join(vaultPath, ".obsidian", "plugins", pluginId);
  }
}

if (!target) {
  console.error(
    "Could not find Obsidian vault. Set OBSIDIAN_VAULT environment variable."
  );
  process.exit(1);
}

console.log(`Copying plugin to: ${target}`);
cpSync(pluginDir, target, { recursive: true, force: true });
console.log("Done. Reload Obsidian or enable the plugin to see changes.");
