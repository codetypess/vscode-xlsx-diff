import { getResolvedDisplayLanguage } from "../display-language";
import {
    formatI18nMessage,
    getRuntimeMessagesForLanguage,
    type I18nLanguage,
    type RuntimeMessages,
} from "./catalog";

export {
    RUNTIME_MESSAGES,
    formatI18nMessage,
    getRuntimeMessagesForLanguage,
    type DiffPanelStrings,
    type I18nLanguage,
    type RuntimeMessages,
} from "./catalog";

export function getRuntimeMessages(
    language: I18nLanguage = getResolvedDisplayLanguage()
): RuntimeMessages {
    return getRuntimeMessagesForLanguage(language);
}
