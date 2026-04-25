import * as vscode from "vscode";
import { gitWorkbookResourceProvider } from "../git/resource-info";
import {
    svnTreeWorkbookResourceProvider,
    svnWorkbookResourceProvider,
} from "../svn/resource-info";

export interface WorkbookResourceDetail {
    label: string;
    value: string;
    titleValue?: string;
    extraFacts?: WorkbookResourceFact[];
}

export interface WorkbookResourceFact {
    label: string;
    value: string;
    titleValue?: string;
}

export interface WorkbookDiffUris {
    original: vscode.Uri;
    modified: vscode.Uri;
}

export interface ScmWorkbookResourceInfo {
    readonly provider: string;
    readonly uri: vscode.Uri;
    readonly resourcePath: string;
    readonly displayPath?: string;
    readonly comparisonPaths?: readonly string[];
    readonly ref?: string;
}

export interface ScmWorkbookResourceProvider {
    readonly scheme: string;
    getResourceInfo(uri: vscode.Uri): ScmWorkbookResourceInfo | undefined;
    getResourceTimeLabel?(info: ScmWorkbookResourceInfo): string | undefined;
    getResourceDetail?(
        info: ScmWorkbookResourceInfo
    ): Promise<WorkbookResourceDetail | undefined>;
    normalizeDiffUris?(diffUris: WorkbookDiffUris): WorkbookDiffUris;
    readWorkbookArchive?(info: ScmWorkbookResourceInfo): Promise<Uint8Array | undefined>;
    isEmptyWorkbook?(info: ScmWorkbookResourceInfo): boolean;
}

const scmWorkbookResourceProviders: readonly ScmWorkbookResourceProvider[] = [
    gitWorkbookResourceProvider,
    svnWorkbookResourceProvider,
    svnTreeWorkbookResourceProvider,
];

function getProvider(providerName: string): ScmWorkbookResourceProvider | undefined {
    return scmWorkbookResourceProviders.find((provider) => provider.scheme === providerName);
}

export function hasScmWorkbookResourceProvider(scheme: string): boolean {
    return scmWorkbookResourceProviders.some((provider) => provider.scheme === scheme);
}

export function getScmWorkbookResourceInfo(
    uri: vscode.Uri
): ScmWorkbookResourceInfo | undefined {
    for (const provider of scmWorkbookResourceProviders) {
        if (uri.scheme !== provider.scheme) {
            continue;
        }

        const info = provider.getResourceInfo(uri);
        if (info) {
            return info;
        }
    }

    return undefined;
}

export function getScmWorkbookResourceTimeLabel(
    info: ScmWorkbookResourceInfo
): string | undefined {
    return getProvider(info.provider)?.getResourceTimeLabel?.(info);
}

export async function getScmWorkbookResourceDetail(
    info: ScmWorkbookResourceInfo
): Promise<WorkbookResourceDetail | undefined> {
    return getProvider(info.provider)?.getResourceDetail?.(info);
}

export async function readScmWorkbookArchive(
    info: ScmWorkbookResourceInfo
): Promise<Uint8Array | undefined> {
    return getProvider(info.provider)?.readWorkbookArchive?.(info);
}

export function isEmptyScmWorkbook(info: ScmWorkbookResourceInfo): boolean {
    return getProvider(info.provider)?.isEmptyWorkbook?.(info) ?? false;
}

export function normalizeScmWorkbookDiffUris(diffUris: WorkbookDiffUris): WorkbookDiffUris {
    let normalizedDiffUris = diffUris;

    for (const provider of scmWorkbookResourceProviders) {
        normalizedDiffUris = provider.normalizeDiffUris?.(normalizedDiffUris) ?? normalizedDiffUris;
    }

    return normalizedDiffUris;
}
