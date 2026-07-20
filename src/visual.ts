"use strict";

import powerbi from "powerbi-visuals-api";
import * as d3 from "d3";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as valueFormatter from "powerbi-visuals-utils-formattingutils/lib/src/valueFormatter";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;

import { VisualFormattingSettingsModel } from "./settings";

const SVG_NS = "http://www.w3.org/2000/svg";
// Hard safety cap: without this, a small manual bucket width relative to the data range
// (or a large manual bucket count) can generate an enormous number of buckets and freeze the browser,
// since bucket computation is O(buckets x rows).
const MAX_BUCKETS = 500;

/**
 * Extracts the literal prefix/suffix text and decimal precision directly from a Power BI number
 * format string's positive section (e.g. "€ #,0.00;€ -#,0.00 ;€ #,0.00" -> prefix "€ ", 2 decimals).
 * Deliberately not using the powerbi-visuals-utils-formattingutils valueFormatter here - it proved
 * unreliable across several real dynamic/multi-section format strings (wrong scale unit, dropped
 * decimals, dropped currency symbol). Manual parsing is fully deterministic.
 */
function parseFormatSection(section: string): { prefix: string; suffix: string } {
    const match = /[#0][#0,. ]*[#0]|[#0]/.exec(section || "");
    if (!match) {
        return { prefix: "", suffix: "" };
    }
    const numPattern = match[0];
    return {
        prefix: section.slice(0, match.index),
        suffix: section.slice(match.index + numPattern.length)
    };
}

/** Plain-loop min/max: Math.min(...array)/Math.max(...array) throws "Maximum call stack size
 * exceeded" on large arrays (a real JS engine limit, not a Power BI one) - this never breaks. */
function safeMin(values: number[]): number {
    let m = Infinity;
    for (const v of values) { if (v < m) m = v; }
    return m;
}
function safeMax(values: number[]): number {
    let m = -Infinity;
    for (const v of values) { if (v > m) m = v; }
    return m;
}


type ChartType = "ring" | "bar" | "column" | "line";
type BucketMode = "manual" | "auto" | "quantile";

interface CategoryPoint {
    catValue: number;
    measureValue: number;
    tooltipValues: number[];
    selectionId: ISelectionId;
}

interface Bucket {
    edgeLow: number;
    edgeHigh: number;
    label: string;
    value: number;
    color: string;
    selectionIds: ISelectionId[];
    extraTooltips: number[];
}

function hexToRgb(hex: string): [number, number, number] {
    const clean = (hex || "#000000").replace("#", "");
    const full = clean.length === 3 ? clean.split("").map(c => c + c).join("") : clean;
    const num = parseInt(full, 16);
    return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
}
function rgbToHex(r: number, g: number, b: number): string {
    const c = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
    return `#${c(r)}${c(g)}${c(b)}`;
}
function lerpColor(a: string, b: string, t: number): string {
    const ca = hexToRgb(a), cb = hexToRgb(b);
    return rgbToHex(ca[0] + (cb[0] - ca[0]) * t, ca[1] + (cb[1] - ca[1]) * t, ca[2] + (cb[2] - ca[2]) * t);
}

let measureCanvas: HTMLCanvasElement | null = null;
function measureTextWidth(text: string, fontSize: number, fontFamily: string): number {
    if (!measureCanvas) {
        measureCanvas = document.createElement("canvas");
    }
    const ctx = measureCanvas.getContext("2d");
    if (!ctx) {
        return text.length * fontSize * 0.55;
    }
    ctx.font = `${fontSize}px ${fontFamily}`;
    return ctx.measureText(text).width;
}

/** Word-wraps text to fit within maxWidth, up to maxLines lines. The last line gets an ellipsis
 * if there's still text left over after filling every allowed line. */
function wrapTextLines(text: string, maxWidth: number, fontSize: number, maxLines: number): string[] {
    const words = (text || "").split(/\s+/).filter(w => w.length > 0);
    if (words.length === 0) return [""];

    const lines: string[] = [];
    let current = "";
    let i = 0;
    while (i < words.length) {
        const candidate = current ? `${current} ${words[i]}` : words[i];
        if (!current || measureTextWidth(candidate, fontSize, "Segoe UI, sans-serif") <= maxWidth) {
            current = candidate;
            i++;
        } else if (lines.length + 1 >= Math.max(1, maxLines)) {
            const rest = `${current} ${words.slice(i).join(" ")}`;
            lines.push(truncateWithEllipsis(rest, maxWidth, fontSize));
            return lines;
        } else {
            lines.push(current);
            current = "";
        }
    }
    if (current) lines.push(current);
    return lines.slice(0, Math.max(1, maxLines));
}

/** Truncates text with a trailing "…" so it fits within maxWidth, without wrapping. */
function truncateWithEllipsis(text: string, maxWidth: number, fontSize: number): string {
    if (measureTextWidth(text, fontSize, "Segoe UI, sans-serif") <= maxWidth) return text;
    let lo = 0, hi = text.length;
    while (lo < hi) {
        const mid = Math.ceil((lo + hi) / 2);
        const candidate = text.slice(0, mid) + "…";
        if (measureTextWidth(candidate, fontSize, "Segoe UI, sans-serif") <= maxWidth) lo = mid;
        else hi = mid - 1;
    }
    return lo > 0 ? text.slice(0, lo) + "…" : "…";
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private tooltipColumnNames: string[] = [];
    private currentValueFormat: string = "";
    private currentBucketTitle: string = "Bucket";
    private currentValueTitle: string = "Value";
    private svg: SVGSVGElement;
    private switcherEl: HTMLDivElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService(options.host.createLocalizationManager());
        this.target = options.element;
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.target.style.overflow = "hidden";
        this.target.style.position = "relative";

        this.switcherEl = document.createElement("div");
        this.switcherEl.className = "bucketChartSwitcher";
        this.target.appendChild(this.switcherEl);

        this.svg = document.createElementNS(SVG_NS, "svg") as SVGSVGElement;
        this.svg.setAttribute("class", "bucketChartSvg");
        this.target.appendChild(this.svg);

        // Click on empty background clears the selection (standard cross-filter UX)
        this.svg.addEventListener("click", (e) => {
            if (e.target === this.svg) {
                this.selectionManager.clear();
                this.applySelectionStyles();
            }
        });

        // Right-click on empty background: whole-visual context menu (empty selection ID)
        this.svg.addEventListener("contextmenu", (e: MouseEvent) => {
            if (e.target !== this.svg) {
                return;
            }
            e.preventDefault();
            if (this.host.hostCapabilities.allowInteractions === false) {
                return;
            }
            this.selectionManager.showContextMenu(
                this.host.createSelectionIdBuilder().createSelectionId(),
                { x: e.clientX, y: e.clientY }
            );
        });
    }

    public update(options: VisualUpdateOptions) {
        this.host.eventService.renderingStarted(options);
        try {
            this.render(options);
            this.host.eventService.renderingFinished(options);
        } catch (err) {
            this.host.eventService.renderingFailed(options, String(err));
            throw err;
        }
    }

    private render(options: VisualUpdateOptions) {
        const dataView: DataView = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

        while (this.svg.firstChild) {
            this.svg.removeChild(this.svg.firstChild);
        }
        while (this.switcherEl.firstChild) {
            this.switcherEl.removeChild(this.switcherEl.firstChild);
        }

        const width = options.viewport.width;
        const height = options.viewport.height;

        const hasData = !!(dataView && dataView.categorical && dataView.categorical.categories &&
            dataView.categorical.categories[0] && dataView.categorical.values && dataView.categorical.values[0]);

        if (!hasData) {
            this.svg.setAttribute("width", String(width));
            this.svg.setAttribute("height", String(height));
            const msg = document.createElementNS(SVG_NS, "text");
            msg.setAttribute("x", String(width / 2));
            msg.setAttribute("y", String(height / 2));
            msg.setAttribute("text-anchor", "middle");
            msg.setAttribute("font-family", "Segoe UI, sans-serif");
            msg.setAttribute("font-size", "12");
            msg.setAttribute("fill", "#605E5C");
            msg.textContent = "Add an Identity field, a Bucket by metric, and a Values metric to get started.";
            this.svg.appendChild(msg);
            return;
        }

        const categoryCol: DataViewCategoryColumn = dataView.categorical.categories[0];
        const allValueCols = dataView.categorical.values;
        const bucketByCol = allValueCols.filter(c => c.source.roles && c.source.roles["bucketBy"])[0];
        const valueCol = allValueCols.filter(c => c.source.roles && c.source.roles["values"])[0];
        const tooltipCols = allValueCols.filter(c => c.source.roles && c.source.roles["tooltips"]);
        this.tooltipColumnNames = tooltipCols.map(c => c.source.displayName);

        if (!valueCol) {
            return;
        }

        // Detect which of the 3 modes applies:
        //  A) "Bucket by" is bound -> bucket that metric's values
        //  B) "Bucket by" empty + Dimension is numeric -> bucket the Dimension's own values
        //  C) "Bucket by" empty + Dimension is text -> no binning at all, one bucket per distinct value
        const firstCategoryValue = categoryCol.values.find(v => v !== null && v !== undefined);
        const isTextMode = !bucketByCol && typeof firstCategoryValue !== "number";

        if (isTextMode) {
            this.renderTextMode(categoryCol, valueCol, tooltipCols, options);
            return;
        }

        const points: CategoryPoint[] = [];
        for (let i = 0; i < categoryCol.values.length; i++) {
            const bv = bucketByCol ? bucketByCol.values[i] : categoryCol.values[i];
            const mv = valueCol.values[i];
            if (typeof bv === "number" && typeof mv === "number") {
                const selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(categoryCol, i)
                    .createSelectionId();
                const tooltipValues = tooltipCols.map(tc => {
                    const v = tc.values[i];
                    return typeof v === "number" ? v : 0;
                });
                points.push({ catValue: bv, measureValue: mv, tooltipValues, selectionId });
            }
        }
        points.sort((a, b) => a.catValue - b.catValue);

        if (points.length === 0) {
            return;
        }

        const bucketModeSettings = this.formattingSettings.bucketsCard;
        const chartTypeSettings = this.formattingSettings.appearanceCard;

        const activeBucketMode = this.resolveBucketMode(bucketModeSettings);
        const activeChartType = this.resolveChartType(chartTypeSettings);

        this.renderSwitchers(bucketModeSettings, chartTypeSettings, activeBucketMode, activeChartType);

        const switcherPosition = chartTypeSettings.switcherPosition.value.value as string; // "top" | "bottom"
        const switcherAlignment = chartTypeSettings.switcherAlignment.value.value as string; // "left" | "center" | "right"
        this.switcherEl.style.justifyContent = switcherAlignment === "center" ? "center" : switcherAlignment === "right" ? "flex-end" : "flex-start";
        this.switcherEl.querySelectorAll<HTMLDivElement>(".bucketChartSwitcherRow").forEach(row => {
            row.style.justifyContent = switcherAlignment === "center" ? "center" : switcherAlignment === "right" ? "flex-end" : "flex-start";
        });

        const switcherHeight = this.switcherEl.childElementCount > 0 ? Math.ceil(this.switcherEl.getBoundingClientRect().height) || 28 : 0;
        const chartWidth = width;
        const chartHeight = Math.max(10, height - switcherHeight);

        if (switcherPosition === "bottom") {
            this.switcherEl.style.position = "absolute";
            this.switcherEl.style.top = "";
            this.switcherEl.style.bottom = "0px";
            this.svg.style.top = "0px";
        } else {
            this.switcherEl.style.position = "relative";
            this.switcherEl.style.bottom = "";
            this.svg.style.top = `${switcherHeight}px`;
        }

        this.svg.setAttribute("width", String(chartWidth));
        this.svg.setAttribute("height", String(chartHeight));
        this.svg.style.position = "absolute";
        this.svg.style.left = "0px";

        const edges = this.computeEdges(points, activeBucketMode);
        let buckets = this.computeBuckets(points, edges, bucketByCol ? bucketByCol.source.format : categoryCol.source.format);
        if (!this.formattingSettings.bucketsCard.showEmptyBuckets.value) {
            buckets = buckets.filter(b => b.value !== 0);
        }
        buckets = this.sortBuckets(buckets);
        if (buckets.length === 0) {
            return;
        }

        const colorsSettings = this.formattingSettings.appearanceCard;
        const isHighContrast = !!(this.host.colorPalette && this.host.colorPalette.isHighContrast);
        const hcForeground = isHighContrast ? this.host.colorPalette.foreground.value : null;
        const maxBucketValue = d3.max(buckets, b => b.value) || 1;

        buckets.forEach((b, i) => {
            if (isHighContrast) {
                b.color = hcForeground;
            } else if (colorsSettings.useCustomColors.value) {
                b.color = colorsSettings.bucketColors[i % colorsSettings.bucketColors.length].value.value;
            } else {
                const t = buckets.length > 1 ? i / (buckets.length - 1) : 0;
                b.color = lerpColor(colorsSettings.startColor.value.value, colorsSettings.endColor.value.value, t);
            }
        });

        const valueObjectsArray = (valueCol.objects || []) as any[];
        const valueDynamicFormat = valueObjectsArray
            .map(o => o && o.general && o.general.formatString)
            .find(f => !!f);
        const valueFormat = valueDynamicFormat || valueCol.source.format;
        this.currentValueFormat = valueFormat;
        this.currentValueTitle = valueCol.source.displayName;
        this.currentBucketTitle = this.formattingSettings.bucketsCard.titleText.value || (bucketByCol ? bucketByCol.source.displayName : categoryCol.source.displayName);

        if (activeChartType === "ring") {
            this.renderRing(buckets, chartWidth, chartHeight, valueFormat, (bucketByCol ? bucketByCol.source.displayName : categoryCol.source.displayName));
        } else if (activeChartType === "line") {
            this.renderLineOrBars(buckets, chartWidth, chartHeight, valueFormat, "line", (bucketByCol ? bucketByCol.source.displayName : categoryCol.source.displayName));
        } else if (activeChartType === "bar") {
            this.renderHorizontalBars(buckets, chartWidth, chartHeight, valueFormat, (bucketByCol ? bucketByCol.source.displayName : categoryCol.source.displayName));
        } else {
            this.renderLineOrBars(buckets, chartWidth, chartHeight, valueFormat, "column", (bucketByCol ? bucketByCol.source.displayName : categoryCol.source.displayName));
        }

        this.applySelectionStyles();
    }

    /** Mode C: Dimension is text and no "Bucket by" metric is bound - no binning happens at all.
     * Each distinct Dimension value is already its own aggregated row (Power BI grouped it via the
     * category role), so we just wrap each row directly as a one-row "bucket" and skip the whole
     * edges/binning pipeline and the bucket-mode switcher (it has nothing to do in this mode). */
    private renderTextMode(categoryCol: DataViewCategoryColumn, valueCol: DataViewValueColumn, tooltipCols: DataViewValueColumn[], options: VisualUpdateOptions) {
        const width = options.viewport.width;
        const height = options.viewport.height;

        let buckets: Bucket[] = [];
        for (let i = 0; i < categoryCol.values.length; i++) {
            const mv = valueCol.values[i];
            if (typeof mv !== "number") continue;
            const selectionId = this.host.createSelectionIdBuilder().withCategory(categoryCol, i).createSelectionId();
            const extraTooltips = tooltipCols.map(tc => (typeof tc.values[i] === "number" ? (tc.values[i] as number) : 0));
            buckets.push({
                edgeLow: i, edgeHigh: i + 1,
                label: String(categoryCol.values[i]),
                value: mv, color: "#CCCCCC",
                selectionIds: [selectionId], extraTooltips
            });
        }
        if (!this.formattingSettings.bucketsCard.showEmptyBuckets.value) {
            buckets = buckets.filter(b => b.value !== 0);
        }
        buckets = this.sortBuckets(buckets);
        if (buckets.length === 0) return;

        const chartTypeSettings = this.formattingSettings.appearanceCard;
        const bucketModeSettings = this.formattingSettings.bucketsCard;
        const activeChartType = this.resolveChartType(chartTypeSettings);

        // No bucket-mode switcher in this mode - binning doesn't apply, so "Manual/Automatic/..." would do nothing.
        this.renderSwitchers(bucketModeSettings, chartTypeSettings, "manual", activeChartType, false);

        const switcherPosition = chartTypeSettings.switcherPosition.value.value as string;
        const switcherAlignment = chartTypeSettings.switcherAlignment.value.value as string;
        this.switcherEl.style.justifyContent = switcherAlignment === "center" ? "center" : switcherAlignment === "right" ? "flex-end" : "flex-start";
        this.switcherEl.querySelectorAll<HTMLDivElement>(".bucketChartSwitcherRow").forEach(row => {
            row.style.justifyContent = switcherAlignment === "center" ? "center" : switcherAlignment === "right" ? "flex-end" : "flex-start";
        });

        const switcherHeight = this.switcherEl.childElementCount > 0 ? Math.ceil(this.switcherEl.getBoundingClientRect().height) || 28 : 0;
        const chartWidth = width;
        const chartHeight = Math.max(10, height - switcherHeight);

        if (switcherPosition === "bottom") {
            this.switcherEl.style.position = "absolute";
            this.switcherEl.style.top = "";
            this.switcherEl.style.bottom = "0px";
            this.svg.style.top = "0px";
        } else {
            this.switcherEl.style.position = "relative";
            this.switcherEl.style.bottom = "";
            this.svg.style.top = `${switcherHeight}px`;
        }

        this.svg.setAttribute("width", String(chartWidth));
        this.svg.setAttribute("height", String(chartHeight));
        this.svg.style.position = "absolute";
        this.svg.style.left = "0px";

        const colorsSettings = this.formattingSettings.appearanceCard;
        const isHighContrast = !!(this.host.colorPalette && this.host.colorPalette.isHighContrast);
        const hcForeground = isHighContrast ? this.host.colorPalette.foreground.value : null;

        buckets.forEach((b, i) => {
            if (isHighContrast) {
                b.color = hcForeground;
            } else if (colorsSettings.useCustomColors.value) {
                b.color = colorsSettings.bucketColors[i % colorsSettings.bucketColors.length].value.value;
            } else {
                const t = buckets.length > 1 ? i / (buckets.length - 1) : 0;
                b.color = lerpColor(colorsSettings.startColor.value.value, colorsSettings.endColor.value.value, t);
            }
        });

        const valueObjectsArray = (valueCol.objects || []) as any[];
        const valueDynamicFormat = valueObjectsArray.map(o => o && o.general && o.general.formatString).find(f => !!f);
        const valueFormat = valueDynamicFormat || valueCol.source.format;
        this.currentValueFormat = valueFormat;
        this.currentValueTitle = valueCol.source.displayName;
        this.currentBucketTitle = this.formattingSettings.bucketsCard.titleText.value || categoryCol.source.displayName;

        if (activeChartType === "ring") {
            this.renderRing(buckets, chartWidth, chartHeight, valueFormat, categoryCol.source.displayName);
        } else if (activeChartType === "line") {
            this.renderLineOrBars(buckets, chartWidth, chartHeight, valueFormat, "line", categoryCol.source.displayName);
        } else if (activeChartType === "bar") {
            this.renderHorizontalBars(buckets, chartWidth, chartHeight, valueFormat, categoryCol.source.displayName);
        } else {
            this.renderLineOrBars(buckets, chartWidth, chartHeight, valueFormat, "column", categoryCol.source.displayName);
        }

        this.applySelectionStyles();
    }

    private resolveBucketMode(s: VisualFormattingSettingsModel["bucketsCard"]): BucketMode {
        const persisted = s.selectedMode.value as BucketMode;
        const allowed: BucketMode[] = [];
        if (s.allowManual.value) allowed.push("manual");
        if (s.allowAuto.value) allowed.push("auto");
        if (s.allowQuantile.value) allowed.push("quantile");

        if (persisted && allowed.indexOf(persisted) >= 0) {
            return persisted;
        }
        const def = s.defaultMode.value.value as BucketMode;
        return allowed.length === 0 || allowed.indexOf(def) >= 0 ? def : allowed[0];
    }

    private resolveChartType(s: VisualFormattingSettingsModel["appearanceCard"]): ChartType {
        const persisted = s.selectedType.value as ChartType;
        const allowed: ChartType[] = [];
        if (s.allowRing.value) allowed.push("ring");
        if (s.allowBar.value) allowed.push("bar");
        if (s.allowColumn.value) allowed.push("column");
        if (s.allowLine.value) allowed.push("line");

        if (persisted && allowed.indexOf(persisted) >= 0) {
            return persisted;
        }
        const def = s.defaultType.value.value as ChartType;
        return allowed.length === 0 || allowed.indexOf(def) >= 0 ? def : allowed[0];
    }

    private renderSwitchers(
        bucketModeSettings: VisualFormattingSettingsModel["bucketsCard"],
        chartTypeSettings: VisualFormattingSettingsModel["appearanceCard"],
        activeBucketMode: BucketMode,
        activeChartType: ChartType,
        showBucketModeSwitcher: boolean = true
    ) {
        const style = this.formattingSettings.appearanceCard;

        const makeRow = (): HTMLDivElement => {
            const row = document.createElement("div");
            row.className = "bucketChartSwitcherRow";
            return row;
        };

        const makeButton = (label: string, active: boolean, onClick: () => void): HTMLDivElement => {
            const btn = document.createElement("div");
            btn.className = "bucketChartSwitcherBtn";
            btn.textContent = label;
            btn.style.fontSize = `${style.switcherFontSize.value}px`;
            btn.style.background = active ? style.switcherActiveColor.value.value : style.switcherBackgroundColor.value.value;
            btn.style.color = active ? style.switcherActiveTextColor.value.value : style.switcherTextColor.value.value;
            btn.addEventListener("click", onClick);
            return btn;
        };

        const makeNumberBox = (labelText: string, value: number, onCommit: (v: number) => void): HTMLDivElement => {
            const wrap = document.createElement("div");
            wrap.className = "bucketChartSwitcherNumberBox";
            wrap.style.fontSize = `${style.switcherFontSize.value}px`;
            wrap.style.color = style.switcherTextColor.value.value;
            wrap.style.background = style.switcherBackgroundColor.value.value;

            const label = document.createElement("span");
            label.textContent = labelText;
            wrap.appendChild(label);

            const input = document.createElement("input");
            input.type = "number";
            input.value = String(value);
            input.className = "bucketChartSwitcherNumberInput";
            input.addEventListener("change", () => {
                const parsed = parseFloat(input.value);
                if (!isNaN(parsed)) {
                    onCommit(parsed);
                }
            });
            wrap.appendChild(input);
            return wrap;
        };

        if (showBucketModeSwitcher) {
            const row = makeRow();
            const options_: Array<[BucketMode, string]> = [];
            if (bucketModeSettings.allowAuto.value) options_.push(["auto", "Automatic"]);
            if (bucketModeSettings.allowQuantile.value) options_.push(["quantile", "Similar dist."]);
            if (bucketModeSettings.allowManual.value) options_.push(["manual", "Manual"]);

            // Only show the mode buttons when there's actually a choice to make - one enabled
            // option means the mode is effectively fixed, so a switcher would do nothing.
            if (options_.length > 1) {
                options_.forEach(([mode, label]) => {
                    row.appendChild(makeButton(label, mode === activeBucketMode, () => {
                        this.host.persistProperties({
                            merge: [{ objectName: "buckets", selector: undefined, properties: { selectedMode: mode } }]
                        });
                    }));
                });
            }

            // Live-adjust box: only when the current mode/sub-mode makes a single numeric knob meaningful
            if (activeBucketMode === "manual" && bucketModeSettings.inputMode.value.value === "width") {
                row.appendChild(makeNumberBox("Width", bucketModeSettings.bucketWidth.value, (v) => {
                    this.host.persistProperties({ merge: [{ objectName: "buckets", selector: undefined, properties: { bucketWidth: v } }] });
                }));
            } else if (activeBucketMode === "auto" && bucketModeSettings.countMode.value.value === "manual") {
                row.appendChild(makeNumberBox("Buckets", bucketModeSettings.autoBucketCount.value, (v) => {
                    this.host.persistProperties({ merge: [{ objectName: "buckets", selector: undefined, properties: { autoBucketCount: v } }] });
                }));
            }

            if (row.childElementCount > 0) {
                this.switcherEl.appendChild(row);
            }
        }

        {
            const row = makeRow();
            const options_: Array<[ChartType, string]> = [];
            if (chartTypeSettings.allowRing.value) options_.push(["ring", "Ring"]);
            if (chartTypeSettings.allowBar.value) options_.push(["bar", "Bar"]);
            if (chartTypeSettings.allowColumn.value) options_.push(["column", "Column"]);
            if (chartTypeSettings.allowLine.value) options_.push(["line", "Line"]);

            // Only show the type buttons when there's actually a choice to make.
            if (options_.length > 1) {
                options_.forEach(([type, label]) => {
                    row.appendChild(makeButton(label, type === activeChartType, () => {
                        this.host.persistProperties({
                            merge: [{ objectName: "appearance", selector: undefined, properties: { selectedType: type } }]
                        });
                    }));
                });
            }
            if (row.childElementCount > 0) {
                this.switcherEl.appendChild(row);
            }
        }
    }

    private computeEdges(points: CategoryPoint[], mode: BucketMode): number[] {
        const values = points.map(p => p.catValue);
        const dataMin = safeMin(values);
        const dataMax = safeMax(values);

        if (mode === "manual") {
            const manual = this.formattingSettings.bucketsCard;
            if (manual.inputMode.value.value === "edges") {
                const parsed = (manual.customEdges.value || "")
                    .split(",")
                    .map(s => parseFloat(s.trim()))
                    .filter(n => !isNaN(n))
                    .sort((a, b) => a - b)
                    .slice(0, MAX_BUCKETS + 1);
                if (parsed.length >= 2) {
                    return parsed;
                }
            }
            const width = Math.max(0.0001, manual.bucketWidth.value);
            const start = manual.autoStart.value ? dataMin : Math.min(manual.bucketStart.value, dataMin);
            const end = manual.autoEnd.value ? dataMax : Math.max(manual.bucketEnd.value, start);
            const rawCount = Math.ceil((end - start) / width) + 1;
            const bucketCount = Math.max(1, Math.min(MAX_BUCKETS, rawCount));
            const edges: number[] = [];
            for (let i = 0; i <= bucketCount; i++) {
                edges.push(start + i * width);
            }
            return edges;
        }

        if (mode === "quantile") {
            const bucketCount = Math.max(1, Math.min(MAX_BUCKETS, Math.round(this.formattingSettings.bucketsCard.quantileBucketCount.value)));
            const sorted = values.slice().sort((a, b) => a - b);
            const edges: number[] = [sorted[0]];
            for (let i = 1; i < bucketCount; i++) {
                const idx = Math.min(sorted.length - 1, Math.floor((i / bucketCount) * sorted.length));
                edges.push(sorted[idx]);
            }
            edges.push(sorted[sorted.length - 1] + 1e-9);
            return Array.from(new Set(edges)).sort((a, b) => a - b);
        }

        const auto = this.formattingSettings.bucketsCard;
        let bucketCount: number;
        if (auto.countMode.value.value === "manual") {
            bucketCount = Math.max(1, Math.min(MAX_BUCKETS, Math.round(auto.autoBucketCount.value)));
        } else {
            bucketCount = Math.max(1, Math.min(MAX_BUCKETS, Math.ceil(Math.log2(points.length) + 1)));
        }
        const span = (dataMax - dataMin) || 1;
        const width = span / bucketCount;
        const edges: number[] = [];
        for (let i = 0; i <= bucketCount; i++) {
            edges.push(dataMin + i * width);
        }
        return edges;
    }

    private sortBuckets(buckets: Bucket[]): Bucket[] {
        const sortBy = this.formattingSettings.bucketsCard.sortBy.value.value as string;
        const sorted = buckets.slice();
        if (sortBy === "rangeAsc") sorted.sort((a, b) => a.edgeLow - b.edgeLow);
        else if (sortBy === "rangeDesc") sorted.sort((a, b) => b.edgeLow - a.edgeLow);
        else if (sortBy === "valueAsc") sorted.sort((a, b) => a.value - b.value);
        else if (sortBy === "valueDesc") sorted.sort((a, b) => b.value - a.value);
        return sorted;
    }

    private computeBuckets(points: CategoryPoint[], edges: number[], dimFormat: string): Bucket[] {
        const aggMethod = this.formattingSettings.valuesCard.aggregationMethod.value.value as string;
        const formatEdge = (n: number) => this.formatBucketEdge(n);

        const buckets: Bucket[] = [];
        for (let i = 0; i < edges.length - 1; i++) {
            const low = edges[i];
            const high = edges[i + 1];
            const isLast = i === edges.length - 2;
            const inBucket = points.filter(p => p.catValue >= low && (isLast ? p.catValue <= high : p.catValue < high));

            let value = 0;
            if (inBucket.length > 0) {
                const vals = inBucket.map(p => p.measureValue);
                if (aggMethod === "average") value = vals.reduce((a, b) => a + b, 0) / vals.length;
                else if (aggMethod === "count") value = inBucket.length;
                else if (aggMethod === "min") value = safeMin(vals);
                else if (aggMethod === "max") value = safeMax(vals);
                else value = vals.reduce((a, b) => a + b, 0);
            }

            const tooltipCount = points.length > 0 ? points[0].tooltipValues.length : 0;
            const extraTooltips: number[] = [];
            for (let t = 0; t < tooltipCount; t++) {
                extraTooltips.push(inBucket.reduce((sum, p) => sum + (p.tooltipValues[t] || 0), 0));
            }

            buckets.push({
                edgeLow: low,
                edgeHigh: high,
                label: `${formatEdge(low)} - ${formatEdge(high)}`,
                value,
                color: "#CCCCCC",
                selectionIds: inBucket.map(p => p.selectionId),
                extraTooltips
            });
        }
        return buckets.filter(b => b.edgeHigh > b.edgeLow);
    }

    /** Click handling shared by all chart types: bucket click = select all underlying rows (cross-filters the report); background click = clear. */
    private attachInteraction(el: SVGElement, bucket: Bucket, tooltipItems: VisualTooltipDataItem[]) {
        el.style.cursor = "pointer";
        (el as any).__bucket = bucket;

        el.setAttribute("tabindex", "0");
        el.setAttribute("role", "button");
        el.setAttribute("aria-label", `${bucket.label}: ${this.formatValue(bucket.value)}`);

        const activate = (e: Event) => {
            e.stopPropagation();
            const multiSelect = (e as MouseEvent).ctrlKey || (e as MouseEvent).metaKey;
            const currentIds = this.selectionManager.getSelectionIds() as ISelectionId[];
            const isSameSingleSelection = !multiSelect && currentIds.length === bucket.selectionIds.length &&
                bucket.selectionIds.every(id => currentIds.some(cid => cid.equals(id)));

            if (isSameSingleSelection) {
                this.selectionManager.clear().then(() => this.applySelectionStyles());
            } else {
                this.selectionManager.select(bucket.selectionIds, multiSelect).then(() => this.applySelectionStyles());
            }
        };

        el.addEventListener("click", activate);
        el.addEventListener("keydown", (e: KeyboardEvent) => {
            if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                activate(e);
            }
        });

        el.addEventListener("contextmenu", (e: MouseEvent) => {
            e.preventDefault();
            e.stopPropagation();
            if (this.host.hostCapabilities.allowInteractions === false) {
                return;
            }
            const contextId = bucket.selectionIds[0] || this.host.createSelectionIdBuilder().createSelectionId();
            this.selectionManager.showContextMenu(contextId, { x: e.clientX, y: e.clientY });
        });

        el.addEventListener("mouseover", (e: MouseEvent) => {
            if (this.host.hostCapabilities.allowInteractions === false) {
                return;
            }
            this.host.tooltipService.show({
                dataItems: tooltipItems,
                identities: bucket.selectionIds,
                coordinates: [e.clientX, e.clientY],
                isTouchEvent: false
            });
        });
        el.addEventListener("mousemove", (e: MouseEvent) => {
            if (this.host.hostCapabilities.allowInteractions === false) {
                return;
            }
            this.host.tooltipService.move({
                dataItems: tooltipItems,
                identities: bucket.selectionIds,
                coordinates: [e.clientX, e.clientY],
                isTouchEvent: false
            });
        });
        el.addEventListener("mouseleave", () => {
            this.host.tooltipService.hide({ isTouchEvent: false, immediately: true });
        });
    }

    /** Dims shapes belonging to buckets not part of the current selection (standard cross-filter visual feedback). */
    private applySelectionStyles() {
        const selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];
        const hasSelection = selectedIds.length > 0;

        this.svg.querySelectorAll<SVGElement>("[data-bucket-shape]").forEach(el => {
            const bucket: Bucket = (el as any).__bucket;
            if (!bucket) return;
            const isSelected = !hasSelection || bucket.selectionIds.some(id => selectedIds.some(sid => sid.equals(id)));
            el.setAttribute("opacity", isSelected ? "1" : "0.3");
        });
    }

    private buildTooltip(bucket: Bucket, valueFormat: string): VisualTooltipDataItem[] {
        const items: VisualTooltipDataItem[] = [
            { displayName: this.currentBucketTitle, value: bucket.label },
            { displayName: this.currentValueTitle, value: this.formatExact(bucket.value) }
        ];
        bucket.extraTooltips.forEach((v, i) => {
            items.push({ displayName: this.tooltipColumnNames[i] || `Metric ${i + 1}`, value: this.formatExact(v) });
        });
        return items;
    }

    /**
     * Display labels round and scale values (K/M/B, fixed decimals) for readability, which can make
     * genuinely different bucket totals look identical. The tooltip always shows the full, unscaled
     * number so the exact figure is never hidden by rounding.
     */
    private formatExact(value: number): string {
        const sections = (this.currentValueFormat || "").split(";");
        const { prefix, suffix } = parseFormatSection(sections[0]);
        const numberText = value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        return `${prefix}${numberText}${suffix}`;
    }

    private formatValue(value: number): string {
        const valuesSettings = this.formattingSettings.valuesCard;
        const decimals = Math.max(0, Math.round(valuesSettings.decimalPlaces.value));
        const unit = valuesSettings.displayUnits.value.value as string;
        const abs = Math.abs(value);

        const sections = (this.currentValueFormat || "").split(";");
        const { prefix, suffix } = parseFormatSection(sections[0]);

        let divisor = 1;
        let unitLetter = "";
        if (unit === "thousands") { divisor = 1e3; unitLetter = "K"; }
        else if (unit === "millions") { divisor = 1e6; unitLetter = "M"; }
        else if (unit === "billions") { divisor = 1e9; unitLetter = "B"; }
        else if (unit === "auto") {
            if (abs >= 1e9) { divisor = 1e9; unitLetter = "B"; }
            else if (abs >= 1e6) { divisor = 1e6; unitLetter = "M"; }
            else if (abs >= 1e3) { divisor = 1e3; unitLetter = "K"; }
        }

        const scaled = value / divisor;
        const numberText = scaled.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
        return `${prefix}${numberText}${unitLetter}${suffix}`;
    }

    private formatBucketEdge(value: number): string {
        const bucketsSettings = this.formattingSettings.bucketsCard;
        const decimals = Math.max(0, Math.round(bucketsSettings.decimalPlaces.value));
        const formatType = bucketsSettings.formatType.value.value as string;

        if (formatType === "percentage") {
            return (value * 100).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + "%";
        }

        const unit = bucketsSettings.displayUnits.value.value as string;
        const abs = Math.abs(value);
        let divisor = 1;
        let suffix = "";
        if (unit === "thousands") { divisor = 1e3; suffix = "K"; }
        else if (unit === "millions") { divisor = 1e6; suffix = "M"; }
        else if (unit === "billions") { divisor = 1e9; suffix = "B"; }
        else if (unit === "auto") {
            if (abs >= 1e9) { divisor = 1e9; suffix = "B"; }
            else if (abs >= 1e6) { divisor = 1e6; suffix = "M"; }
            else if (abs >= 1e3) { divisor = 1e3; suffix = "K"; }
        }

        const scaled = value / divisor;
        return scaled.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + suffix;
    }

    /** Draws a semi-transparent rounded background behind a block of label text lines, so light
     * label colors (tuned for colored wedges) stay legible over plain white chart backgrounds. */
    private renderLabelBackground(centerX: number, centerY: number, texts: string[], fontSize: number, anchor: "middle" | "start" | "end" = "middle") {
        const labelsSettings = this.formattingSettings.labelsCard;
        if (!labelsSettings.showBackground.value || texts.length === 0) return;

        const lineHeight = fontSize * 1.15;
        const maxWidth = Math.max(...texts.map(t => measureTextWidth(t, fontSize, "Segoe UI, sans-serif")));
        const padX = 5;
        const padY = 3;
        const width = maxWidth + padX * 2;
        const height = texts.length * lineHeight + padY * 2;

        let x = centerX - width / 2;
        if (anchor === "start") x = centerX - padX;
        if (anchor === "end") x = centerX - width + padX;

        const rect = document.createElementNS(SVG_NS, "rect");
        rect.setAttribute("x", String(x));
        rect.setAttribute("y", String(centerY - height / 2));
        rect.setAttribute("width", String(width));
        rect.setAttribute("height", String(height));
        rect.setAttribute("rx", "3");
        rect.setAttribute("fill", labelsSettings.backgroundColor.value.value);
        rect.setAttribute("fill-opacity", String(1 - Math.max(0, Math.min(100, labelsSettings.backgroundTransparency.value)) / 100));
        rect.style.pointerEvents = "none";
        this.svg.appendChild(rect);
    }

    private renderRing(buckets: Bucket[], width: number, height: number, valueFormat: string, bucketTitle: string) {
        const legendSettings = this.formattingSettings.legendCard;
        const labelsSettings = this.formattingSettings.labelsCard;
        const valuesSettings = this.formattingSettings.valuesCard;
        const ringSettings = this.formattingSettings.appearanceCard;
        const showLegend = legendSettings.show.value;
        const legendPos = legendSettings.position.value.value as string;

        const isVerticalLegend = legendPos.startsWith("left") || legendPos.startsWith("right");
        const legendSpace = showLegend ? (isVerticalLegend ? (legendSettings.maxWidth.value + 26) : 34) : 0;

        const plotWidth = isVerticalLegend ? width - legendSpace : width;
        const plotHeight = isVerticalLegend ? height : height - legendSpace;
        const legendOffsetX = legendPos.startsWith("left") ? legendSpace : 0;
        const legendTopOffset = legendPos === "top" || legendPos === "topCenter" ? legendSpace : 0;

        const cx = legendOffsetX + plotWidth / 2;
        const cy = legendTopOffset + plotHeight / 2;
        const radius = Math.max(10, Math.min(plotWidth, plotHeight) / 2 - 10);
        const innerRadius = radius * Math.max(0, Math.min(90, ringSettings.ringInnerRadiusPercent.value)) / 100;

        const total = d3.sum(buckets, b => b.value) || 1;
        const pieGen = d3.pie<Bucket>().value(b => Math.max(0, b.value)).sort(null);
        const arcs = pieGen(buckets);
        const arcGen = d3.arc<d3.PieArcDatum<Bucket>>().innerRadius(innerRadius).outerRadius(radius);
        const labelRadius = innerRadius + (radius - innerRadius) * 0.6;
        const labelArcGen = d3.arc<d3.PieArcDatum<Bucket>>().innerRadius(labelRadius).outerRadius(labelRadius);
        const minAngleDeg = labelsSettings.minAngleToShow.value;

        arcs.forEach(a => {
            const angleDeg = ((a.endAngle - a.startAngle) * 180) / Math.PI;

            const group = document.createElementNS(SVG_NS, "g");
            group.setAttribute("data-bucket-shape", "1");

            const path = document.createElementNS(SVG_NS, "path");
            path.setAttribute("d", arcGen(a));
            path.setAttribute("transform", `translate(${cx},${cy})`);
            path.setAttribute("fill", a.data.color);
            path.setAttribute("stroke", "#FFFFFF");
            path.setAttribute("stroke-width", "1");
            group.appendChild(path);
            this.svg.appendChild(group);
            this.attachInteraction(group as unknown as SVGElement, a.data, this.buildTooltip(a.data, valueFormat));

            if (labelsSettings.show.value && angleDeg >= minAngleDeg) {
                const [lx, ly] = labelArcGen.centroid(a);
                const lines: Array<{ text: string; color: string }> = [];
                if (this.formattingSettings.bucketsCard.showRange.value) lines.push({ text: a.data.label, color: this.formattingSettings.bucketsCard.rangeFontColor.value.value });
                const valueParts: string[] = [];
                if (valuesSettings.showValue.value) valueParts.push(this.formatValue(a.data.value));
                if (valuesSettings.showPercent.value) valueParts.push(`${((a.data.value / total) * 100).toFixed(1)}%`);
                if (valueParts.length > 0) lines.push({ text: valueParts.join(" | "), color: valuesSettings.fontColor.value.value });

                if (lines.length > 0) {
                    this.renderLabelBackground(cx + lx, cy + ly, lines.map(l => l.text), labelsSettings.fontSize.value);

                    const text = document.createElementNS(SVG_NS, "text");
                    text.setAttribute("x", String(cx + lx));
                    text.setAttribute("text-anchor", "middle");
                    text.setAttribute("font-family", "Segoe UI, sans-serif");
                    text.setAttribute("font-size", String(labelsSettings.fontSize.value));
                    text.style.pointerEvents = "none";
                    const lineHeight = labelsSettings.fontSize.value * 1.15;
                    const startY = cy + ly - ((lines.length - 1) * lineHeight) / 2;
                    lines.forEach((line, li) => {
                        const tspan = document.createElementNS(SVG_NS, "tspan");
                        tspan.setAttribute("x", String(cx + lx));
                        tspan.setAttribute("y", String(startY + li * lineHeight));
                        tspan.setAttribute("fill", line.color);
                        tspan.textContent = line.text;
                        text.appendChild(tspan);
                    });
                    this.svg.appendChild(text);
                }
            }
        });

        if (showLegend) {
            this.renderLegend(buckets, width, height, legendPos, bucketTitle);
        }
    }

    private renderLegend(buckets: Bucket[], width: number, height: number, position: string, bucketTitle: string) {
        const legendSettings = this.formattingSettings.legendCard;
        const swatch = 10;
        const fontSize = legendSettings.fontSize.value;
        const rowHeight = fontSize + 6;
        const maxLabelWidth = legendSettings.maxWidth.value;

        const bucketsForTitle = this.formattingSettings.bucketsCard;
        const titleText = bucketsForTitle.showTitle.value ? (bucketsForTitle.titleText.value || bucketTitle) : "";
        const hasTitle = !!titleText;

        const isHorizontal = position.startsWith("top") || position.startsWith("bottom");

        if (isHorizontal) {
            const displayLabels = buckets.map(b => truncateWithEllipsis(b.label, maxLabelWidth, fontSize));
            const titleWidth = hasTitle ? measureTextWidth(titleText + ":", fontSize, "Segoe UI, sans-serif") + 16 : 0;
            const itemWidths = displayLabels.map(l => swatch + 5 + measureTextWidth(l, fontSize, "Segoe UI, sans-serif") + 14);
            const totalWidth = titleWidth + itemWidths.reduce((a, b) => a + b, 0);
            let x = position.endsWith("Center") ? Math.max(8, (width - totalWidth) / 2) : 8;
            const y = position.startsWith("top") ? 14 : height - 12;

            if (hasTitle) {
                const titleEl = document.createElementNS(SVG_NS, "text");
                titleEl.setAttribute("x", String(x));
                titleEl.setAttribute("y", String(y + 4));
                titleEl.setAttribute("font-family", "Segoe UI, sans-serif");
                titleEl.setAttribute("font-size", String(fontSize));
                titleEl.setAttribute("font-weight", "bold");
                titleEl.setAttribute("fill", legendSettings.fontColor.value.value);
                titleEl.textContent = titleText + ":";
                this.svg.appendChild(titleEl);
                x += titleWidth;
            }

            buckets.forEach((b, i) => {
                const rect = document.createElementNS(SVG_NS, "rect");
                rect.setAttribute("x", String(x));
                rect.setAttribute("y", String(y - swatch / 2));
                rect.setAttribute("width", String(swatch));
                rect.setAttribute("height", String(swatch));
                rect.setAttribute("fill", b.color);
                this.svg.appendChild(rect);

                const text = document.createElementNS(SVG_NS, "text");
                text.setAttribute("x", String(x + swatch + 5));
                text.setAttribute("y", String(y + 4));
                text.setAttribute("font-family", "Segoe UI, sans-serif");
                text.setAttribute("font-size", String(fontSize));
                text.setAttribute("fill", legendSettings.fontColor.value.value);
                text.textContent = displayLabels[i];
                this.svg.appendChild(text);

                x += itemWidths[i];
            });
        } else {
            const lineHeight = fontSize * 1.15;
            const entries = buckets.map(b => legendSettings.wrapText.value
                ? wrapTextLines(b.label, maxLabelWidth, fontSize, Math.max(1, legendSettings.maxLines.value))
                : [truncateWithEllipsis(b.label, maxLabelWidth, fontSize)]
            );
            const entryHeights = entries.map(lines => Math.max(rowHeight, lines.length * lineHeight + 4));
            const titleHeight = hasTitle ? fontSize + 8 : 0;
            const totalHeight = titleHeight + entryHeights.reduce((a, b) => a + b, 0);
            const x = position.startsWith("right") ? width - maxLabelWidth - swatch - 12 : 8;
            let y = position.endsWith("Center") ? Math.max(14, (height - totalHeight) / 2) : 14;

            if (hasTitle) {
                const titleEl = document.createElementNS(SVG_NS, "text");
                titleEl.setAttribute("x", String(x));
                titleEl.setAttribute("y", String(y));
                titleEl.setAttribute("font-family", "Segoe UI, sans-serif");
                titleEl.setAttribute("font-size", String(fontSize));
                titleEl.setAttribute("font-weight", "bold");
                titleEl.setAttribute("fill", legendSettings.fontColor.value.value);
                titleEl.textContent = titleText;
                this.svg.appendChild(titleEl);
                y += titleHeight;
            }

            buckets.forEach((b, i) => {
                const lines = entries[i];
                const ry = y + swatch / 2 + 2;
                const rect = document.createElementNS(SVG_NS, "rect");
                rect.setAttribute("x", String(x));
                rect.setAttribute("y", String(ry - swatch / 2));
                rect.setAttribute("width", String(swatch));
                rect.setAttribute("height", String(swatch));
                rect.setAttribute("fill", b.color);
                this.svg.appendChild(rect);

                lines.forEach((line, li) => {
                    const text = document.createElementNS(SVG_NS, "text");
                    text.setAttribute("x", String(x + swatch + 5));
                    text.setAttribute("y", String(ry + 4 + li * lineHeight));
                    text.setAttribute("font-family", "Segoe UI, sans-serif");
                    text.setAttribute("font-size", String(fontSize));
                    text.setAttribute("fill", legendSettings.fontColor.value.value);
                    text.textContent = line;
                    this.svg.appendChild(text);
                });

                y += entryHeights[i];
            });
        }
    }

    private renderHorizontalBars(buckets: Bucket[], width: number, height: number, valueFormat: string, bucketTitle: string) {
        const axisSettings = this.formattingSettings.axisCard;
        const labelsSettings = this.formattingSettings.labelsCard;
        const valuesSettings = this.formattingSettings.valuesCard;

        const axisFontFamily = "Segoe UI, sans-serif";
        const rawMaxLabelWidth = axisSettings.show.value
            ? Math.max(0, ...buckets.map(b => measureTextWidth(b.label, axisSettings.fontSize.value, axisFontFamily)))
            : 0;
        const leftMargin = axisSettings.show.value
            ? Math.min(axisSettings.maxAxisWidth.value, Math.max(40, rawMaxLabelWidth + 10))
            : 10;
        const rightMargin = 20;
        const topMargin = 10;
        const titleHeight = this.formattingSettings.bucketsCard.showTitle.value ? axisSettings.fontSize.value + 8 : 0;
        const bottomMargin = (axisSettings.show.value ? 24 : 8) + titleHeight;

        const plotW = Math.max(10, width - leftMargin - rightMargin);
        const plotH = Math.max(10, height - topMargin - bottomMargin);

        if (this.formattingSettings.bucketsCard.showTitle.value) {
            const titleText = this.formattingSettings.bucketsCard.titleText.value || bucketTitle;
            const titleEl = document.createElementNS(SVG_NS, "text");
            titleEl.setAttribute("x", String(leftMargin + plotW / 2));
            titleEl.setAttribute("y", String(height - titleHeight + axisSettings.fontSize.value));
            titleEl.setAttribute("text-anchor", "middle");
            titleEl.setAttribute("font-family", "Segoe UI, sans-serif");
            titleEl.setAttribute("font-size", String(axisSettings.fontSize.value));
            titleEl.setAttribute("font-weight", "bold");
            titleEl.setAttribute("fill", axisSettings.fontColor.value.value);
            titleEl.textContent = titleText;
            this.svg.appendChild(titleEl);
        }

        const maxVal = d3.max(buckets, b => b.value) || 1;
        const xScale = d3.scaleLinear().domain([0, maxVal]).range([0, plotW]);
        const yScale = d3.scaleBand().domain(buckets.map(b => b.label)).range([0, plotH]).padding(0.2);

        if (axisSettings.showGridlines.value) {
            xScale.ticks(5).forEach(t => {
                const gx = leftMargin + xScale(t);
                const line = document.createElementNS(SVG_NS, "line");
                line.setAttribute("x1", String(gx));
                line.setAttribute("x2", String(gx));
                line.setAttribute("y1", String(topMargin));
                line.setAttribute("y2", String(topMargin + plotH));
                line.setAttribute("stroke", "#E1E1E1");
                this.svg.appendChild(line);
            });
        }

        const total = d3.sum(buckets, b => b.value) || 1;

        buckets.forEach(b => {
            const by = topMargin + (yScale(b.label) || 0);
            const bw = xScale(Math.max(0, b.value));

            const rect = document.createElementNS(SVG_NS, "rect");
            rect.setAttribute("data-bucket-shape", "1");
            rect.setAttribute("x", String(leftMargin));
            rect.setAttribute("y", String(by));
            rect.setAttribute("width", String(bw));
            rect.setAttribute("height", String(yScale.bandwidth()));
            rect.setAttribute("fill", b.color);
            this.svg.appendChild(rect);
            this.attachInteraction(rect, b, this.buildTooltip(b, valueFormat));

            if (axisSettings.show.value) {
                const availableWidth = leftMargin - 6;
                const fontSize = axisSettings.fontSize.value;
                if (axisSettings.wrapText.value) {
                    const lines = wrapTextLines(b.label, availableWidth, fontSize, Math.max(1, axisSettings.maxLines.value));
                    const lineHeight = fontSize * 1.15;
                    const startY = by + yScale.bandwidth() / 2 + 4 - ((lines.length - 1) * lineHeight) / 2;
                    lines.forEach((line, li) => {
                        const catLabel = document.createElementNS(SVG_NS, "text");
                        catLabel.setAttribute("x", String(leftMargin - 6));
                        catLabel.setAttribute("y", String(startY + li * lineHeight));
                        catLabel.setAttribute("text-anchor", "end");
                        catLabel.setAttribute("font-size", String(fontSize));
                        catLabel.setAttribute("fill", axisSettings.fontColor.value.value);
                        catLabel.textContent = line;
                        this.svg.appendChild(catLabel);
                    });
                } else {
                    const catLabel = document.createElementNS(SVG_NS, "text");
                    catLabel.setAttribute("x", String(leftMargin - 6));
                    catLabel.setAttribute("y", String(by + yScale.bandwidth() / 2 + 4));
                    catLabel.setAttribute("text-anchor", "end");
                    catLabel.setAttribute("font-size", String(fontSize));
                    catLabel.setAttribute("fill", axisSettings.fontColor.value.value);
                    catLabel.textContent = truncateWithEllipsis(b.label, availableWidth, fontSize);
                    this.svg.appendChild(catLabel);
                }
            }

            if (labelsSettings.show.value && (valuesSettings.showValue.value || valuesSettings.showPercent.value)) {
                const parts: string[] = [];
                if (valuesSettings.showValue.value) parts.push(this.formatValue(b.value));
                if (valuesSettings.showPercent.value) parts.push(`${((b.value / total) * 100).toFixed(1)}%`);
                const text = parts.join(" | ");

                const fontSize = labelsSettings.fontSize.value;
                const textWidth = measureTextWidth(text, fontSize, "Segoe UI, sans-serif");
                const fitsOutside = leftMargin + bw + 4 + textWidth <= width - 4;

                const centerY = by + yScale.bandwidth() / 2 + 4 - fontSize * 0.35;
                if (fitsOutside) {
                    this.renderLabelBackground(leftMargin + bw + 4, centerY, [text], fontSize, "start");
                } else {
                    this.renderLabelBackground(leftMargin + bw - 4, centerY, [text], fontSize, "end");
                }

                const valLabel = document.createElementNS(SVG_NS, "text");
                valLabel.setAttribute("y", String(by + yScale.bandwidth() / 2 + 4));
                valLabel.setAttribute("font-size", String(fontSize));
                valLabel.setAttribute("fill", valuesSettings.fontColor.value.value);
                valLabel.style.pointerEvents = "none";
                valLabel.textContent = text;

                if (fitsOutside) {
                    valLabel.setAttribute("x", String(leftMargin + bw + 4));
                } else {
                    valLabel.setAttribute("x", String(leftMargin + bw - 4));
                    valLabel.setAttribute("text-anchor", "end");
                }
                this.svg.appendChild(valLabel);
            }
        });
    }

    private renderLineOrBars(buckets: Bucket[], width: number, height: number, valueFormat: string, mode: "line" | "column", bucketTitle: string) {
        const axisSettings = this.formattingSettings.axisCard;
        const labelsSettings = this.formattingSettings.labelsCard;
        const valuesSettings = this.formattingSettings.valuesCard;

        const leftMargin = 50;
        const rightMargin = 15;
        const topMargin = labelsSettings.show.value ? 24 : 10;
        const titleHeight = this.formattingSettings.bucketsCard.showTitle.value ? axisSettings.fontSize.value + 8 : 0;
        const axisFontFamily = "Segoe UI, sans-serif";
        const maxCatLabelWidth = axisSettings.show.value
            ? Math.max(0, ...buckets.map(b => measureTextWidth(b.label, axisSettings.fontSize.value, axisFontFamily)))
            : 0;
        let extraAxisHeight = 0;
        if (axisSettings.show.value && axisSettings.rotateLabels.value) {
            extraAxisHeight = Math.min(110, maxCatLabelWidth * 0.6);
        } else if (axisSettings.show.value && axisSettings.wrapText.value) {
            extraAxisHeight = (Math.max(1, axisSettings.maxLines.value) - 1) * axisSettings.fontSize.value * 1.15;
        }
        const bottomMargin = (axisSettings.show.value ? 40 : 10) + titleHeight + extraAxisHeight;

        const plotW = Math.max(10, width - leftMargin - rightMargin);
        const plotH = Math.max(10, height - topMargin - bottomMargin);

        if (this.formattingSettings.bucketsCard.showTitle.value) {
            const titleText = this.formattingSettings.bucketsCard.titleText.value || bucketTitle;
            const titleEl = document.createElementNS(SVG_NS, "text");
            titleEl.setAttribute("x", String(leftMargin + plotW / 2));
            titleEl.setAttribute("y", String(height - titleHeight + axisSettings.fontSize.value));
            titleEl.setAttribute("text-anchor", "middle");
            titleEl.setAttribute("font-family", "Segoe UI, sans-serif");
            titleEl.setAttribute("font-size", String(axisSettings.fontSize.value));
            titleEl.setAttribute("font-weight", "bold");
            titleEl.setAttribute("fill", axisSettings.fontColor.value.value);
            titleEl.textContent = titleText;
            this.svg.appendChild(titleEl);
        }

        const maxVal = d3.max(buckets, b => b.value) || 1;
        const total = d3.sum(buckets, b => b.value) || 1;
        const yScale = d3.scaleLinear().domain([0, maxVal]).range([plotH, 0]);
        const xScale = d3.scaleBand().domain(buckets.map(b => b.label)).range([0, plotW]).padding(0.25);

        if (axisSettings.showGridlines.value) {
            yScale.ticks(4).forEach(t => {
                const gy = topMargin + yScale(t);
                const line = document.createElementNS(SVG_NS, "line");
                line.setAttribute("x1", String(leftMargin));
                line.setAttribute("x2", String(leftMargin + plotW));
                line.setAttribute("y1", String(gy));
                line.setAttribute("y2", String(gy));
                line.setAttribute("stroke", "#E1E1E1");
                this.svg.appendChild(line);
            });
        }

        if (axisSettings.show.value) {
            const fontSize = axisSettings.fontSize.value;
            const bandWidth = xScale.bandwidth();
            buckets.forEach(b => {
                const bx = leftMargin + (xScale(b.label) || 0) + bandWidth / 2;
                const baseY = topMargin + plotH + 16;

                if (axisSettings.rotateLabels.value) {
                    const label = document.createElementNS(SVG_NS, "text");
                    label.setAttribute("x", String(bx));
                    label.setAttribute("y", String(baseY - 4));
                    label.setAttribute("text-anchor", "end");
                    label.setAttribute("font-size", String(fontSize));
                    label.setAttribute("fill", axisSettings.fontColor.value.value);
                    label.setAttribute("transform", `rotate(-35, ${bx}, ${baseY - 4})`);
                    label.textContent = truncateWithEllipsis(b.label, 140, fontSize);
                    this.svg.appendChild(label);
                } else if (axisSettings.wrapText.value) {
                    const lines = wrapTextLines(b.label, Math.max(20, bandWidth), fontSize, Math.max(1, axisSettings.maxLines.value));
                    const lineHeight = fontSize * 1.15;
                    lines.forEach((line, li) => {
                        const label = document.createElementNS(SVG_NS, "text");
                        label.setAttribute("x", String(bx));
                        label.setAttribute("y", String(baseY + li * lineHeight));
                        label.setAttribute("text-anchor", "middle");
                        label.setAttribute("font-size", String(fontSize));
                        label.setAttribute("fill", axisSettings.fontColor.value.value);
                        label.textContent = line;
                        this.svg.appendChild(label);
                    });
                } else {
                    const label = document.createElementNS(SVG_NS, "text");
                    label.setAttribute("x", String(bx));
                    label.setAttribute("y", String(baseY));
                    label.setAttribute("text-anchor", "middle");
                    label.setAttribute("font-size", String(fontSize));
                    label.setAttribute("fill", axisSettings.fontColor.value.value);
                    label.textContent = truncateWithEllipsis(b.label, Math.max(20, bandWidth), fontSize);
                    this.svg.appendChild(label);
                }
            });
        }

        if (mode === "column") {
            buckets.forEach(b => {
                const bx = leftMargin + (xScale(b.label) || 0);
                const by = topMargin + yScale(Math.max(0, b.value));
                const bh = plotH - yScale(Math.max(0, b.value));

                const rect = document.createElementNS(SVG_NS, "rect");
                rect.setAttribute("data-bucket-shape", "1");
                rect.setAttribute("x", String(bx));
                rect.setAttribute("y", String(by));
                rect.setAttribute("width", String(xScale.bandwidth()));
                rect.setAttribute("height", String(Math.max(0, bh)));
                rect.setAttribute("fill", b.color);
                this.svg.appendChild(rect);
                this.attachInteraction(rect, b, this.buildTooltip(b, valueFormat));

                if (labelsSettings.show.value && (valuesSettings.showValue.value || valuesSettings.showPercent.value)) {
                    const parts: string[] = [];
                    if (valuesSettings.showValue.value) parts.push(this.formatValue(b.value));
                    if (valuesSettings.showPercent.value) parts.push(`${((b.value / total) * 100).toFixed(1)}%`);

                    this.renderLabelBackground(bx + xScale.bandwidth() / 2, by - 4 - labelsSettings.fontSize.value * 0.35, [parts.join(" | ")], labelsSettings.fontSize.value);

                    const valLabel = document.createElementNS(SVG_NS, "text");
                    valLabel.setAttribute("x", String(bx + xScale.bandwidth() / 2));
                    valLabel.setAttribute("y", String(by - 4));
                    valLabel.setAttribute("text-anchor", "middle");
                    valLabel.setAttribute("font-size", String(labelsSettings.fontSize.value));
                    valLabel.setAttribute("fill", valuesSettings.fontColor.value.value);
                    valLabel.style.pointerEvents = "none";
                    valLabel.textContent = parts.join(" | ");
                    this.svg.appendChild(valLabel);
                }
            });
        } else {
            const linePoints = buckets.map(b => ({
                x: leftMargin + (xScale(b.label) || 0) + xScale.bandwidth() / 2,
                y: topMargin + yScale(Math.max(0, b.value))
            }));
            const lineGen = d3.line<{ x: number; y: number }>().x(p => p.x).y(p => p.y);
            const path = document.createElementNS(SVG_NS, "path");
            path.setAttribute("d", lineGen(linePoints));
            path.setAttribute("fill", "none");
            path.setAttribute("stroke", this.formattingSettings.appearanceCard.endColor.value.value);
            path.setAttribute("stroke-width", "2");
            this.svg.appendChild(path);

            linePoints.forEach((p, i) => {
                const dot = document.createElementNS(SVG_NS, "circle");
                dot.setAttribute("data-bucket-shape", "1");
                dot.setAttribute("cx", String(p.x));
                dot.setAttribute("cy", String(p.y));
                dot.setAttribute("r", "5.5");
                dot.setAttribute("fill", buckets[i].color);
                this.svg.appendChild(dot);
                this.attachInteraction(dot, buckets[i], this.buildTooltip(buckets[i], valueFormat));

                if (labelsSettings.show.value && (valuesSettings.showValue.value || valuesSettings.showPercent.value)) {
                    const parts: string[] = [];
                    if (valuesSettings.showValue.value) parts.push(this.formatValue(buckets[i].value));
                    if (valuesSettings.showPercent.value) parts.push(`${((buckets[i].value / total) * 100).toFixed(1)}%`);

                    this.renderLabelBackground(p.x, p.y - 10 - labelsSettings.fontSize.value * 0.35, [parts.join(" | ")], labelsSettings.fontSize.value);

                    const valLabel = document.createElementNS(SVG_NS, "text");
                    valLabel.setAttribute("x", String(p.x));
                    valLabel.setAttribute("y", String(p.y - 10));
                    valLabel.setAttribute("text-anchor", "middle");
                    valLabel.setAttribute("font-size", String(labelsSettings.fontSize.value));
                    valLabel.setAttribute("fill", valuesSettings.fontColor.value.value);
                    valLabel.style.pointerEvents = "none";
                    valLabel.textContent = parts.join(" | ");
                    this.svg.appendChild(valLabel);
                }
            });
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
