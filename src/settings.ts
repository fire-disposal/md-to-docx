import { App, PluginSettingTab, Setting } from "obsidian";
import { t } from "./i18n";
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

    containerEl.createEl("h2", { text: t("pluginName") });

    // ── Export ──
    containerEl.createEl("h3", { text: t("settingsExport") });

    new Setting(containerEl)
      .setName(t("settingsOutputFolder"))
      .setDesc(t("settingsOutputFolderDesc"))
      .addText((text) =>
        text
          .setPlaceholder(t("settingsOutputPlaceholder"))
          .setValue(this.plugin.settings.outputFolder)
          .onChange(async (value) => {
            this.plugin.settings.outputFolder = value.replace(/^\/|\/$/g, "");
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(t("settingsOverwrite"))
      .setDesc(t("settingsOverwriteDesc"))
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.overwriteExisting)
          .onChange(async (value) => {
            this.plugin.settings.overwriteExisting = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(t("settingsOpenAfter"))
      .setDesc(t("settingsOpenAfterDesc"))
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.openAfterExport)
          .onChange(async (value) => {
            this.plugin.settings.openAfterExport = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(t("settingsPageSize"))
      .setDesc(t("settingsPageSizeDesc"))
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

    // ── Images ──
    containerEl.createEl("h3", { text: t("settingsImages") });

    new Setting(containerEl)
      .setName(t("settingsEmbedImages"))
      .setDesc(t("settingsEmbedImagesDesc"))
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.embedImages)
          .onChange(async (value) => {
            this.plugin.settings.embedImages = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(t("settingsCaptions"))
      .setDesc(t("settingsCaptionsDesc"))
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.addCaptions)
          .onChange(async (value) => {
            this.plugin.settings.addCaptions = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(t("settingsMaxWidth"))
      .setDesc(t("settingsMaxWidthDesc"))
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
