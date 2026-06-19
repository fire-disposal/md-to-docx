import { App, PluginSettingTab, Setting } from "obsidian";
import type MdToDocxPlugin from "../main";

export interface MdToDocxSettings {
  overwriteExisting: boolean;
  embedImages: boolean;
  addCaptions: boolean;
  pageSize: "A4" | "Letter";
  outputFolder: string;
  openAfterExport: boolean;
  imageMaxWidth: number;
}

export const DEFAULT_SETTINGS: MdToDocxSettings = {
  overwriteExisting: true,
  embedImages: true,
  addCaptions: true,
  pageSize: "A4",
  outputFolder: "",
  openAfterExport: true,
  imageMaxWidth: 6,
};

export class MdToDocxSettingTab extends PluginSettingTab {
  plugin: MdToDocxPlugin;

  constructor(app: App, plugin: MdToDocxPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "Markdown to DOCX" });

    // --- Export ---
    containerEl.createEl("h3", { text: "Export" });

    new Setting(containerEl)
      .setName("Output folder")
      .setDesc(
        "Save .docx files to a subfolder (e.g. 'exports'). Leave empty to save next to the note."
      )
      .addText((text) =>
        text
          .setPlaceholder("e.g. exports")
          .setValue(this.plugin.settings.outputFolder)
          .onChange(async (value) => {
            this.plugin.settings.outputFolder = value.replace(/^\/|\/$/g, "");
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Overwrite existing files")
      .setDesc("Silently overwrite an existing .docx with the same name.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.overwriteExisting)
          .onChange(async (value) => {
            this.plugin.settings.overwriteExisting = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Open after export")
      .setDesc("Open the .docx file with the system default application after export.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.openAfterExport)
          .onChange(async (value) => {
            this.plugin.settings.openAfterExport = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Page size")
      .setDesc("Document page dimensions.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("A4", "A4")
          .addOption("Letter", "Letter")
          .setValue(this.plugin.settings.pageSize)
          .onChange(async (value) => {
            this.plugin.settings.pageSize = value as "A4" | "Letter";
            await this.plugin.saveSettings();
          })
      );

    // --- Images ---
    containerEl.createEl("h3", { text: "Images" });

    new Setting(containerEl)
      .setName("Embed images")
      .setDesc(
        "Embed referenced local images into the .docx. PNG, JPEG, GIF, BMP, WebP are supported."
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.embedImages)
          .onChange(async (value) => {
            this.plugin.settings.embedImages = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Add image captions")
      .setDesc(
        "Insert the image alt-text as a caption below each embedded image."
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.addCaptions)
          .onChange(async (value) => {
            this.plugin.settings.addCaptions = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Max image width (inches)")
      .setDesc(
        "Scale images down if wider than this. Set to 0 to disable. Default: 6 inches."
      )
      .addSlider((slider) =>
        slider
          .setLimits(0, 12, 0.5)
          .setValue(this.plugin.settings.imageMaxWidth)
          .setDynamicTooltip()
          .onChange(async (value) => {
            this.plugin.settings.imageMaxWidth = value;
            await this.plugin.saveSettings();
          })
      );
  }
}
