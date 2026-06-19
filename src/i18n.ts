import { getLanguage } from "obsidian";

const strings: Record<string, Record<string, string>> = {
  en: {
    pluginName: "MD to DOCX",
    ribbonTooltip: "Export to DOCX",
    cmdExport: "Export current note to .docx",
    cmdExportOverwrite: "Export current note to .docx (overwrite)",
    noticeNoActive: "No active note.",
    noticeNotMd: "Active file is not a Markdown note.",
    noticeExists: "File already exists: ",
    noticeFolderFail: "Failed to create folder: ",
    noticeExported: "Exported to ",
    noticeUpdated: "Updated ",
    noticeExportFail: "Export failed — ",

    settingsExport: "Export",
    settingsOutputFolder: "Output folder",
    settingsOutputFolderDesc:
      "Save .docx files to a subfolder (e.g. 'exports'). Leave empty to save next to the note.",
    settingsOutputPlaceholder: "e.g. exports",
    settingsOverwrite: "Overwrite existing files",
    settingsOverwriteDesc:
      "Silently overwrite an existing .docx with the same name.",
    settingsOpenAfter: "Open after export",
    settingsOpenAfterDesc:
      "Open the .docx with the system default application after export.",
    settingsPageSize: "Page size",
    settingsPageSizeDesc: "Document page dimensions.",
    settingsImages: "Images",
    settingsEmbedImages: "Embed images",
    settingsEmbedImagesDesc:
      "Embed referenced local images into the .docx. PNG, JPEG, GIF, BMP, WebP are supported.",
    settingsCaptions: "Add image captions",
    settingsCaptionsDesc:
      "Insert the image alt-text as a caption below each embedded image.",
    settingsMaxWidth: "Max image width (inches)",
    settingsMaxWidthDesc:
      "Scale images down if wider than this. Set to 0 to disable.",
  },

  zh: {
    pluginName: "MD to DOCX",
    ribbonTooltip: "导出为 DOCX",
    cmdExport: "导出当前笔记为 .docx",
    cmdExportOverwrite: "导出当前笔记为 .docx（覆盖）",
    noticeNoActive: "没有活动的笔记。",
    noticeNotMd: "当前文件不是 Markdown 笔记。",
    noticeExists: "文件已存在：",
    noticeFolderFail: "创建文件夹失败：",
    noticeExported: "已导出到 ",
    noticeUpdated: "已更新 ",
    noticeExportFail: "导出失败 — ",

    settingsExport: "导出",
    settingsOutputFolder: "输出目录",
    settingsOutputFolderDesc:
      "将 .docx 保存到子目录（如 'exports'）。留空则保存到笔记同目录。",
    settingsOutputPlaceholder: "如 exports",
    settingsOverwrite: "自动覆盖同名文件",
    settingsOverwriteDesc: "静默覆盖已存在的同名 .docx 文件。",
    settingsOpenAfter: "导出后打开",
    settingsOpenAfterDesc: "导出后使用系统默认程序打开 .docx 文件。",
    settingsPageSize: "页面尺寸",
    settingsPageSizeDesc: "文档页面尺寸。",
    settingsImages: "图片",
    settingsEmbedImages: "嵌入图片",
    settingsEmbedImagesDesc:
      "将笔记中引用的本地图片嵌入 .docx。支持 PNG、JPEG、GIF、BMP、WebP。",
    settingsCaptions: "添加图片题注",
    settingsCaptionsDesc:
      "在每张嵌入图片下方插入题注（取自图片 alt 文本）。",
    settingsMaxWidth: "图片最大宽度（英寸）",
    settingsMaxWidthDesc:
      "超过此宽度的图片等比缩小。设为 0 则不限制。",
  },
};

export function t(key: string): string {
  let lang = "en";
  try {
    const detected = getLanguage();
    if (detected && strings[detected]) {
      lang = detected;
    } else if (detected?.includes("-")) {
      const base = detected.split("-")[0];
      if (strings[base]) lang = base;
    }
  } catch {
    // keep default
  }
  return strings[lang]?.[key] ?? strings.en[key] ?? key;
}
