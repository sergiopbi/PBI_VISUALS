"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import SimpleCard = formattingSettings.SimpleCard;
import CompositeCard = formattingSettings.CompositeCard;
import Group = formattingSettings.Group;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

const DEFAULT_SLICE_COLORS = [
    "#1F3B5C", "#2E5077", "#3D6591", "#4C7AAB", "#6690BC", "#7FA6CC", "#98BBDB",
    "#2A4B70", "#3A6088", "#5578A0", "#7090B4", "#8CA6C4", "#A8BFD3", "#5B8AB8",
    "#456F94", "#6F9BC2", "#84AECF", "#396286", "#5A82A8", "#719AC0"
];

// ============================================================
// Buckets: Mode / Manual / Automatic / Similar distribution / Range format / Title
// Everything about HOW buckets are built and labeled lives here.
// ============================================================
class BucketsCard extends CompositeCard {
    name: string = "buckets";
    displayName: string = "Buckets";
    displayNameKey: string = "Buckets_Card";

    defaultMode = new formattingSettings.ItemDropdown({
        name: "defaultMode", displayName: "Default mode", displayNameKey: "Buckets_DefaultMode",
        items: [
            { displayName: "Manual", value: "manual" },
            { displayName: "Automatic", value: "auto" },
            { displayName: "Similar distribution", value: "quantile" }
        ],
        value: { displayName: "Manual", value: "manual" }
    });
    allowManual = new formattingSettings.ToggleSwitch({ name: "allowManual", displayName: "Manual", displayNameKey: "Buckets_AllowManual", value: true });
    allowAuto = new formattingSettings.ToggleSwitch({ name: "allowAuto", displayName: "Automatic", displayNameKey: "Buckets_AllowAuto", value: true });
    allowQuantile = new formattingSettings.ToggleSwitch({ name: "allowQuantile", displayName: "Similar distribution", displayNameKey: "Buckets_AllowQuantile", value: true });
    showEmptyBuckets = new formattingSettings.ToggleSwitch({ name: "showEmptyBuckets", displayName: "Show buckets with a value of 0", displayNameKey: "Buckets_ShowEmpty", value: true });
    sortBy = new formattingSettings.ItemDropdown({
        name: "sortBy", displayName: "Sort buckets by", displayNameKey: "Buckets_SortBy",
        items: [
            { displayName: "Range (ascending)", value: "rangeAsc" },
            { displayName: "Range (descending)", value: "rangeDesc" },
            { displayName: "Value (ascending)", value: "valueAsc" },
            { displayName: "Value (descending)", value: "valueDesc" }
        ],
        value: { displayName: "Range (ascending)", value: "rangeAsc" }
    });
    selectedMode = new formattingSettings.TextInput({ name: "selectedMode", displayName: "Selected mode (internal)", displayNameKey: "Buckets_SelectedMode", value: "", placeholder: "" });

    modeGroup = new Group({
        name: "mode", displayName: "Mode", displayNameKey: "Buckets_GroupMode",
        slices: [this.defaultMode, this.allowManual, this.allowAuto, this.allowQuantile, this.showEmptyBuckets, this.sortBy, this.selectedMode]
    });

    inputMode = new formattingSettings.ItemDropdown({
        name: "inputMode", displayName: "Define by", displayNameKey: "Buckets_InputMode",
        items: [
            { displayName: "Equal width", value: "width" },
            { displayName: "Custom edges", value: "edges" }
        ],
        value: { displayName: "Equal width", value: "width" }
    });
    bucketWidth = new formattingSettings.NumUpDown({ name: "bucketWidth", displayName: "Bucket width", displayNameKey: "Buckets_Width", value: 15 });
    autoStart = new formattingSettings.ToggleSwitch({ name: "autoStart", displayName: "Auto start (use Dimension minimum)", displayNameKey: "Buckets_AutoStart", value: true });
    bucketStart = new formattingSettings.NumUpDown({ name: "bucketStart", displayName: "Start value", displayNameKey: "Buckets_Start", value: 0 });
    autoEnd = new formattingSettings.ToggleSwitch({ name: "autoEnd", displayName: "Auto end (use Dimension maximum)", displayNameKey: "Buckets_AutoEnd", value: true });
    bucketEnd = new formattingSettings.NumUpDown({ name: "bucketEnd", displayName: "End value", displayNameKey: "Buckets_End", value: 100 });
    customEdges = new formattingSettings.TextInput({ name: "customEdges", displayName: "Edges (comma-separated)", displayNameKey: "Buckets_Edges", value: "", placeholder: "0,15,30,50,100" });

    manualGroup = new Group({
        name: "manual", displayName: "Manual", displayNameKey: "Buckets_GroupManual",
        slices: [this.inputMode, this.bucketWidth, this.autoStart, this.bucketStart, this.autoEnd, this.bucketEnd, this.customEdges]
    });

    countMode = new formattingSettings.ItemDropdown({
        name: "countMode", displayName: "Number of buckets", displayNameKey: "Buckets_CountMode",
        items: [
            { displayName: "Automatic (Sturges' rule)", value: "sturges" },
            { displayName: "Fixed", value: "manual" }
        ],
        value: { displayName: "Automatic (Sturges' rule)", value: "sturges" }
    });
    autoBucketCount = new formattingSettings.NumUpDown({ name: "autoBucketCount", displayName: "Number of buckets", displayNameKey: "Buckets_AutoCount", value: 6 });

    autoGroup = new Group({
        name: "auto", displayName: "Automatic", displayNameKey: "Buckets_GroupAuto",
        slices: [this.countMode, this.autoBucketCount]
    });

    quantileBucketCount = new formattingSettings.NumUpDown({ name: "quantileBucketCount", displayName: "Number of buckets", displayNameKey: "Buckets_QuantileCount", value: 4 });

    quantileGroup = new Group({
        name: "quantile", displayName: "Similar distribution", displayNameKey: "Buckets_GroupQuantile",
        slices: [this.quantileBucketCount]
    });

    showRange = new formattingSettings.ToggleSwitch({ name: "showRange", displayName: "Show bucket range label", displayNameKey: "Buckets_ShowRange", value: true });
    formatType = new formattingSettings.ItemDropdown({
        name: "formatType", displayName: "Format", displayNameKey: "Buckets_FormatType",
        items: [
            { displayName: "Number", value: "number" },
            { displayName: "Percentage", value: "percentage" }
        ],
        value: { displayName: "Number", value: "number" }
    });
    displayUnits = new formattingSettings.ItemDropdown({
        name: "displayUnits", displayName: "Display units", displayNameKey: "Buckets_DisplayUnits",
        items: [
            { displayName: "Auto", value: "auto" },
            { displayName: "None", value: "none" },
            { displayName: "Thousands", value: "thousands" },
            { displayName: "Millions", value: "millions" },
            { displayName: "Billions", value: "billions" }
        ],
        value: { displayName: "Auto", value: "auto" }
    });
    decimalPlaces = new formattingSettings.NumUpDown({ name: "decimalPlaces", displayName: "Decimal places", displayNameKey: "Buckets_DecimalPlaces", value: 0 });
    rangeFontColor = new formattingSettings.ColorPicker({ name: "rangeFontColor", displayName: "Color", displayNameKey: "Buckets_RangeColor", value: { value: "#FFFFFF" } });

    rangeFormatGroup = new Group({
        name: "rangeFormat", displayName: "Range format", displayNameKey: "Buckets_GroupRangeFormat",
        slices: [this.showRange, this.formatType, this.displayUnits, this.decimalPlaces, this.rangeFontColor]
    });

    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayName: "Show title", displayNameKey: "Buckets_ShowTitle", value: false });
    titleText = new formattingSettings.TextInput({ name: "titleText", displayName: "Title text", displayNameKey: "Buckets_TitleText", value: "", placeholder: "e.g. Days, Age, Years" });

    titleGroup = new Group({
        name: "title", displayName: "Title", displayNameKey: "Buckets_GroupTitle",
        slices: [this.showTitle, this.titleText]
    });

    groups: Array<Group> = [this.modeGroup, this.manualGroup, this.autoGroup, this.quantileGroup, this.rangeFormatGroup, this.titleGroup];

    onPreProcess(): void {
        this.selectedMode.visible = false;
        const useWidth = this.inputMode.value.value === "width";
        this.bucketWidth.visible = useWidth;
        this.autoStart.visible = useWidth;
        this.bucketStart.visible = useWidth && !this.autoStart.value;
        this.autoEnd.visible = useWidth;
        this.bucketEnd.visible = useWidth && !this.autoEnd.value;
        this.customEdges.visible = !useWidth;
        this.autoBucketCount.visible = this.countMode.value.value === "manual";
        this.displayUnits.visible = this.formatType.value.value === "number";
    }
}

// ============================================================
// Values: Aggregation / Format / Label
// Everything about the aggregated VALUE (not the bucket range) lives here.
// ============================================================
class ValuesCard extends CompositeCard {
    name: string = "values";
    displayName: string = "Values";
    displayNameKey: string = "Values_Card";

    aggregationMethod = new formattingSettings.ItemDropdown({
        name: "aggregationMethod", displayName: "Summarize by", displayNameKey: "Values_Aggregation",
        items: [
            { displayName: "Sum", value: "sum" },
            { displayName: "Average", value: "average" },
            { displayName: "Count", value: "count" },
            { displayName: "Minimum", value: "min" },
            { displayName: "Maximum", value: "max" }
        ],
        value: { displayName: "Sum", value: "sum" }
    });

    aggregationGroup = new Group({
        name: "aggregation", displayName: "Aggregation", displayNameKey: "Values_GroupAggregation",
        slices: [this.aggregationMethod]
    });

    displayUnits = new formattingSettings.ItemDropdown({
        name: "displayUnits", displayName: "Display units", displayNameKey: "Values_DisplayUnits",
        items: [
            { displayName: "Auto", value: "auto" },
            { displayName: "None", value: "none" },
            { displayName: "Thousands", value: "thousands" },
            { displayName: "Millions", value: "millions" },
            { displayName: "Billions", value: "billions" }
        ],
        value: { displayName: "Auto", value: "auto" }
    });
    decimalPlaces = new formattingSettings.NumUpDown({ name: "decimalPlaces", displayName: "Decimal places", displayNameKey: "Values_DecimalPlaces", value: 0 });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Color", displayNameKey: "Values_FontColor", value: { value: "#EDEDED" } });

    formatGroup = new Group({
        name: "format", displayName: "Format", displayNameKey: "Values_GroupFormat",
        slices: [this.displayUnits, this.decimalPlaces, this.fontColor]
    });

    showValue = new formattingSettings.ToggleSwitch({ name: "showValue", displayName: "Show value label", displayNameKey: "Values_ShowValue", value: true });
    showPercent = new formattingSettings.ToggleSwitch({ name: "showPercent", displayName: "Show percentage of total", displayNameKey: "Values_ShowPercent", value: false });

    labelGroup = new Group({
        name: "label", displayName: "Label", displayNameKey: "Values_GroupLabel",
        slices: [this.showValue, this.showPercent]
    });

    groups: Array<Group> = [this.aggregationGroup, this.formatGroup, this.labelGroup];

    onPreProcess(): void {
        this.displayUnits.visible = true;
    }
}

// ============================================================
// Labels: master show/hide + shared text size + ring label-collision control.
// ============================================================
class LabelsCardSettings extends SimpleCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show labels", displayNameKey: "Labels_Show", value: true });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", displayNameKey: "Labels_FontSize", value: 11 });
    minAngleToShow = new formattingSettings.NumUpDown({ name: "minAngleToShow", displayName: "Min. slice angle to show label (°, ring only)", displayNameKey: "Labels_MinAngle", value: 12 });
    showBackground = new formattingSettings.ToggleSwitch({ name: "showBackground", displayName: "Show background", displayNameKey: "Labels_ShowBackground", value: false });
    backgroundColor = new formattingSettings.ColorPicker({ name: "backgroundColor", displayName: "Background color", displayNameKey: "Labels_BackgroundColor", value: { value: "#FFFFFF" } });
    backgroundTransparency = new formattingSettings.NumUpDown({ name: "backgroundTransparency", displayName: "Background transparency (%)", displayNameKey: "Labels_BackgroundTransparency", value: 20 });

    name: string = "labels";
    displayName: string = "Labels";
    displayNameKey: string = "Labels_Card";
    slices: Array<FormattingSettingsSlice> = [this.show, this.fontSize, this.minAngleToShow, this.showBackground, this.backgroundColor, this.backgroundTransparency];
}

// ============================================================
// Appearance: Chart type / Switcher position / Switcher style / Ring / Colors
// Everything about how the whole visual LOOKS lives here.
// ============================================================
class AppearanceCard extends CompositeCard {
    name: string = "appearance";
    displayName: string = "Appearance";
    displayNameKey: string = "Appearance_Card";

    defaultType = new formattingSettings.ItemDropdown({
        name: "defaultType", displayName: "Default chart type", displayNameKey: "Appearance_DefaultType",
        items: [
            { displayName: "Ring", value: "ring" },
            { displayName: "Bar", value: "bar" },
            { displayName: "Column", value: "column" },
            { displayName: "Line", value: "line" }
        ],
        value: { displayName: "Ring", value: "ring" }
    });
    allowRing = new formattingSettings.ToggleSwitch({ name: "allowRing", displayName: "Ring", displayNameKey: "Appearance_AllowRing", value: true });
    allowBar = new formattingSettings.ToggleSwitch({ name: "allowBar", displayName: "Bar", displayNameKey: "Appearance_AllowBar", value: true });
    allowColumn = new formattingSettings.ToggleSwitch({ name: "allowColumn", displayName: "Column", displayNameKey: "Appearance_AllowColumn", value: true });
    allowLine = new formattingSettings.ToggleSwitch({ name: "allowLine", displayName: "Line", displayNameKey: "Appearance_AllowLine", value: true });
    selectedType = new formattingSettings.TextInput({ name: "selectedType", displayName: "Selected type (internal)", displayNameKey: "Appearance_SelectedType", value: "", placeholder: "" });

    chartTypeGroup = new Group({
        name: "chartType", displayName: "Chart type", displayNameKey: "Appearance_GroupChartType",
        slices: [this.defaultType, this.allowRing, this.allowBar, this.allowColumn, this.allowLine, this.selectedType]
    });

    switcherPosition = new formattingSettings.ItemDropdown({
        name: "switcherPosition", displayName: "Position", displayNameKey: "Switcher_Position",
        items: [{ displayName: "Top", value: "top" }, { displayName: "Bottom", value: "bottom" }],
        value: { displayName: "Top", value: "top" }
    });
    switcherAlignment = new formattingSettings.ItemDropdown({
        name: "switcherAlignment", displayName: "Alignment", displayNameKey: "Switcher_Alignment",
        items: [
            { displayName: "Left", value: "left" },
            { displayName: "Center", value: "center" },
            { displayName: "Right", value: "right" }
        ],
        value: { displayName: "Left", value: "left" }
    });

    switcherPositionGroup = new Group({
        name: "switcherPosition", displayName: "Switcher position", displayNameKey: "Appearance_GroupSwitcherPosition",
        slices: [this.switcherPosition, this.switcherAlignment]
    });

    switcherBackgroundColor = new formattingSettings.ColorPicker({ name: "switcherBackgroundColor", displayName: "Background color", displayNameKey: "Switcher_BackgroundColor", value: { value: "#F3F2F1" } });
    switcherActiveColor = new formattingSettings.ColorPicker({ name: "switcherActiveColor", displayName: "Selected background color", displayNameKey: "Switcher_ActiveColor", value: { value: "#1F3B5C" } });
    switcherTextColor = new formattingSettings.ColorPicker({ name: "switcherTextColor", displayName: "Text color", displayNameKey: "Switcher_TextColor", value: { value: "#605E5C" } });
    switcherActiveTextColor = new formattingSettings.ColorPicker({ name: "switcherActiveTextColor", displayName: "Selected text color", displayNameKey: "Switcher_ActiveTextColor", value: { value: "#FFFFFF" } });
    switcherFontSize = new formattingSettings.NumUpDown({ name: "switcherFontSize", displayName: "Text size", displayNameKey: "Switcher_FontSize", value: 10 });

    switcherStyleGroup = new Group({
        name: "switcherStyle", displayName: "Switcher style", displayNameKey: "Appearance_GroupSwitcherStyle",
        slices: [this.switcherBackgroundColor, this.switcherActiveColor, this.switcherTextColor, this.switcherActiveTextColor, this.switcherFontSize]
    });

    ringInnerRadiusPercent = new formattingSettings.NumUpDown({ name: "ringInnerRadiusPercent", displayName: "Hole size (%)", displayNameKey: "Appearance_RingHole", value: 55 });

    ringGroup = new Group({
        name: "ring", displayName: "Ring", displayNameKey: "Appearance_GroupRing",
        slices: [this.ringInnerRadiusPercent]
    });

    useCustomColors = new formattingSettings.ToggleSwitch({ name: "useCustomColors", displayName: "Custom color per bucket", displayNameKey: "Colors_UseCustomColors", value: false });
    startColor = new formattingSettings.ColorPicker({ name: "startColor", displayName: "Gradient - first bucket", displayNameKey: "Colors_StartColor", value: { value: "#C9D6E3" } });
    endColor = new formattingSettings.ColorPicker({ name: "endColor", displayName: "Gradient - last bucket", displayNameKey: "Colors_EndColor", value: { value: "#1F3B5C" } });
    bucketColors: formattingSettings.ColorPicker[] = Array.from({ length: 20 }, (_, i) =>
        new formattingSettings.ColorPicker({
            name: `bucket${i + 1}`, displayName: `Bucket ${i + 1}`, displayNameKey: `Colors_Bucket${i + 1}`,
            value: { value: DEFAULT_SLICE_COLORS[i % DEFAULT_SLICE_COLORS.length] }
        })
    );

    colorsGroup = new Group({
        name: "colors", displayName: "Colors", displayNameKey: "Appearance_GroupColors",
        slices: [this.useCustomColors, this.startColor, this.endColor, ...this.bucketColors]
    });

    groups: Array<Group> = [this.chartTypeGroup, this.switcherPositionGroup, this.switcherStyleGroup, this.ringGroup, this.colorsGroup];

    onPreProcess(): void {
        this.selectedType.visible = false;
        const custom = this.useCustomColors.value;
        this.startColor.visible = !custom;
        this.endColor.visible = !custom;
        this.bucketColors.forEach(c => { c.visible = custom; });
    }
}

// ============================================================
// Legend
// ============================================================
class LegendCardSettings extends SimpleCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show legend", displayNameKey: "Legend_Show", value: true });
    position = new formattingSettings.ItemDropdown({
        name: "position", displayName: "Position", displayNameKey: "Legend_Position",
        items: [
            { displayName: "Top", value: "top" },
            { displayName: "Top center", value: "topCenter" },
            { displayName: "Bottom", value: "bottom" },
            { displayName: "Bottom center", value: "bottomCenter" },
            { displayName: "Left", value: "left" },
            { displayName: "Left center", value: "leftCenter" },
            { displayName: "Right", value: "right" },
            { displayName: "Right center", value: "rightCenter" }
        ],
        value: { displayName: "Right", value: "right" }
    });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", displayNameKey: "Legend_FontSize", value: 11 });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Text color", displayNameKey: "Legend_FontColor", value: { value: "#252423" } });
    maxWidth = new formattingSettings.NumUpDown({ name: "maxWidth", displayName: "Max label width (px)", displayNameKey: "Legend_MaxWidth", value: 110 });
    wrapText = new formattingSettings.ToggleSwitch({ name: "wrapText", displayName: "Wrap labels", displayNameKey: "Legend_WrapText", value: false });
    maxLines = new formattingSettings.NumUpDown({ name: "maxLines", displayName: "Max lines when wrapping", displayNameKey: "Legend_MaxLines", value: 2 });

    name: string = "legend";
    displayName: string = "Legend";
    displayNameKey: string = "Legend_Card";
    slices: Array<FormattingSettingsSlice> = [this.show, this.position, this.fontSize, this.fontColor, this.maxWidth, this.wrapText, this.maxLines];

    onPreProcess(): void {
        this.maxLines.visible = this.wrapText.value;
    }
}

// ============================================================
// Axis
// ============================================================
class AxisCardSettings extends SimpleCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", displayNameKey: "Axis_Show", value: true });
    showGridlines = new formattingSettings.ToggleSwitch({ name: "showGridlines", displayName: "Show gridlines", displayNameKey: "Axis_ShowGridlines", value: true });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", displayNameKey: "Axis_FontSize", value: 10 });
    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Text color", displayNameKey: "Axis_FontColor", value: { value: "#605E5C" } });
    maxAxisWidth = new formattingSettings.NumUpDown({ name: "maxAxisWidth", displayName: "Max category axis width (px, Bar only)", displayNameKey: "Axis_MaxAxisWidth", value: 120 });
    wrapText = new formattingSettings.ToggleSwitch({ name: "wrapText", displayName: "Wrap category labels", displayNameKey: "Axis_WrapText", value: false });
    maxLines = new formattingSettings.NumUpDown({ name: "maxLines", displayName: "Max lines when wrapping", displayNameKey: "Axis_MaxLines", value: 2 });
    rotateLabels = new formattingSettings.ToggleSwitch({ name: "rotateLabels", displayName: "Rotate category labels (Column/Line only)", displayNameKey: "Axis_RotateLabels", value: false });

    name: string = "axis";
    displayName: string = "Axis";
    displayNameKey: string = "Axis_Card";
    slices: Array<FormattingSettingsSlice> = [
        this.show, this.showGridlines, this.fontSize, this.fontColor,
        this.maxAxisWidth, this.wrapText, this.maxLines, this.rotateLabels
    ];

    onPreProcess(): void {
        this.maxLines.visible = this.wrapText.value;
    }
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    bucketsCard = new BucketsCard();
    valuesCard = new ValuesCard();
    labelsCard = new LabelsCardSettings();
    appearanceCard = new AppearanceCard();
    legendCard = new LegendCardSettings();
    axisCard = new AxisCardSettings();

    cards = [
        this.bucketsCard, this.valuesCard, this.labelsCard,
        this.appearanceCard, this.legendCard, this.axisCard
    ];
}
