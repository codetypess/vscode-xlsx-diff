import * as vscode from "vscode";
import type { I18nLanguage } from "./i18n/catalog";

const CONFIGURATION_SECTION = "xlsx-diff";
const DISPLAY_LANGUAGE_SETTING = "displayLanguage";

export const DISPLAY_LANGUAGE_CONFIGURATION_KEY = `${CONFIGURATION_SECTION}.${DISPLAY_LANGUAGE_SETTING}`;

export type DisplayLanguageSetting = "auto" | "en" | "zh-cn";
export type ResolvedDisplayLanguage = I18nLanguage;

export function resolveDisplayLanguage(
    configuredLanguage: DisplayLanguageSetting | undefined,
    vscodeLanguage: string
): ResolvedDisplayLanguage {
    switch (configuredLanguage) {
        case "en":
            return "en";
        case "zh-cn":
            return "zh-cn";
        case "auto":
        case undefined:
        default:
            return vscodeLanguage.toLowerCase().startsWith("zh") ? "zh-cn" : "en";
    }
}

export function getResolvedDisplayLanguage(): ResolvedDisplayLanguage {
    const configuredLanguage = vscode.workspace
        .getConfiguration(CONFIGURATION_SECTION)
        .get<DisplayLanguageSetting>(DISPLAY_LANGUAGE_SETTING, "auto");

    return resolveDisplayLanguage(configuredLanguage, vscode.env.language);
}

export function isChineseDisplayLanguage(): boolean {
    return getResolvedDisplayLanguage() === "zh-cn";
}

export function getHtmlLanguageTag(): string {
    return getResolvedDisplayLanguage() === "zh-cn" ? "zh-CN" : "en";
}

export function affectsDisplayLanguage(event: vscode.ConfigurationChangeEvent): boolean {
    return event.affectsConfiguration(DISPLAY_LANGUAGE_CONFIGURATION_KEY);
}
