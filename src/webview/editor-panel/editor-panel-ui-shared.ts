import type { EditorSearchResultMessage } from "./editor-panel-types";

export interface PendingSummary {
    sheetKeys: Set<string>;
    rows: Set<number>;
    columns: Set<number>;
}

export type SearchPanelMode = "find" | "replace";
type SearchPanelFeedbackStatus = EditorSearchResultMessage["status"] | "replaced" | "no-change";

export interface SearchPanelFeedback {
    status: SearchPanelFeedbackStatus;
    message?: string;
}

export interface SearchPanelPosition {
    left: number;
    top: number;
}

export type ContextMenuState =
    | {
          kind: "tab";
          sheetKey: string;
          x: number;
          y: number;
      }
    | {
          kind: "row";
          rowNumber: number;
          x: number;
          y: number;
      }
    | {
          kind: "column";
          columnNumber: number;
          x: number;
          y: number;
      };

export function classNames(values: Array<string | false | null | undefined>): string {
    return values.filter(Boolean).join(" ");
}
