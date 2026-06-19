import { Notice, Plugin, TFile } from "obsidian";
import { markdownToDocx, PAGE_A4, PAGE_LETTER } from "./src/docx";
import { resolveImages } from "./src/images";
import { t } from "./src/i18n";
import {
  MdToDocxSettingTab,
  MdToDocxSettings,
  DEFAULT_SETTINGS,
} from "./src/settings";

export default class MdToDocxPlugin extends Plugin {
  settings: MdToDocxSettings = DEFAULT_SETTINGS;

  async onload() {
    await this.loadSettings();

    this.addRibbonIcon("file-down", t("ribbonTooltip"), () => {
      this.exportCurrentNote();
    });

    this.addCommand({
      id: "export-current-note-to-docx",
      name: t("cmdExport"),
      callback: () => this.exportCurrentNote(),
    });

    this.addCommand({
      id: "export-current-note-to-docx-overwrite",
      name: t("cmdExportOverwrite"),
      callback: () => this.exportCurrentNote(true),
    });

    this.addSettingTab(new MdToDocxSettingTab(this.app, this));
  }

  onunload() {}

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  async exportCurrentNote(forceOverwrite = false) {
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      new Notice(t("noticeNoActive"));
      return;
    }

    if (file.extension !== "md") {
      new Notice(t("noticeNotMd"));
      return;
    }

    const outPath = this.getOutputPath(file);

    const folderPath = outPath.substring(0, outPath.lastIndexOf("/"));
    if (folderPath && !(await this.app.vault.adapter.exists(folderPath))) {
      try {
        await this.app.vault.createFolder(folderPath);
      } catch {
        new Notice(t("noticeFolderFail") + folderPath);
        return;
      }
    }

    if (!forceOverwrite && !this.settings.overwriteExisting) {
      if (await this.app.vault.adapter.exists(outPath)) {
        new Notice(t("noticeExists") + outPath);
        return;
      }
    }

    try {
      const markdown = await this.app.vault.read(file);

      let images:
        | Map<string, import("./src/docx/types").ImageBinary>
        | undefined;
      if (this.settings.embedImages) {
        images = await resolveImages(markdown, file, this.app);
      }

      const page =
        this.settings.pageSize === "Letter" ? PAGE_LETTER : PAGE_A4;

      const docxBuffer = markdownToDocx(markdown, {
        images,
        addCaptions: this.settings.addCaptions && this.settings.embedImages,
        page,
        imageMaxWidth: this.settings.imageMaxWidth,
      });

      if (forceOverwrite || this.settings.overwriteExisting) {
        if (await this.app.vault.adapter.exists(outPath)) {
          const existing = this.app.vault.getAbstractFileByPath(outPath);
          if (existing instanceof TFile) {
            await this.app.vault.modifyBinary(existing, docxBuffer);
            new Notice(t("noticeUpdated") + outPath);
            if (this.settings.openAfterExport) {
              this.openWithDefaultApp(outPath);
            }
            return;
          }
        }
      }

      await this.app.vault.createBinary(outPath, docxBuffer);
      new Notice(t("noticeExported") + outPath);

      if (this.settings.openAfterExport) {
        this.openWithDefaultApp(outPath);
      }
    } catch (e) {
      new Notice(t("noticeExportFail") + e);
    }
  }

  private getOutputPath(file: TFile): string {
    const base = file.path.replace(/\.md$/, ".docx");
    const name = base.substring(base.lastIndexOf("/") + 1);
    if (this.settings.outputFolder) {
      return `${this.settings.outputFolder}/${name}`;
    }
    const parent = file.parent?.path;
    if (parent) {
      return `${parent}/${name}`;
    }
    return name;
  }

  private openWithDefaultApp(vaultPath: string) {
    try {
      const adapter = this.app.vault.adapter as any;
      const fullPath = adapter.getFullPath
        ? adapter.getFullPath(vaultPath)
        : `${adapter.basePath}/${vaultPath}`;
      const { shell } = require("electron");
      shell.openPath(fullPath);
    } catch {
      // Electron shell unavailable — not a desktop environment
    }
  }
}
