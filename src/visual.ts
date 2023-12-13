"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

import { VisualFormattingSettingsModel } from "./settings";

import ISelectionId = powerbi.visuals.ISelectionId;
import DataViewTableRow = powerbi.DataViewTableRow;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn

import FilterAction = powerbi.FilterAction;

import {
    SEARCH_DEBOUNCE,
    SELECT_DEBOUNCE,
    FILTER_LIMIT,
} from "./settings";

// powerbi models
import {
    IFilterColumnTarget,
    IAdvancedFilterCondition,
    AdvancedFilter,
    AdvancedFilterLogicalOperators
} from "powerbi-models";

import debounce from "lodash.debounce";

interface searchDataPoint {
    value: DataViewTableRow;
    category?: string;
    color?: string;
    strokeColor?: string;
    strokeWidth?: number;
    selection: ISelectionId;
    format?: string;
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private host: powerbi.extensibility.visual.IVisualHost;
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private dataPoints: searchDataPoint[] = [];

    private searchDebounce: number = SEARCH_DEBOUNCE;
    private selectDebounce: number = SELECT_DEBOUNCE;

    private filterLimit: number = FILTER_LIMIT;

    constructor(options: VisualConstructorOptions) {

        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;

        if (document) {

            const search_input: HTMLElement = document.createElement("input");
            const new_div: HTMLElement = document.createElement("div");
            new_div.classList.add("InputDiv")
            new_div.appendChild(search_input)
            this.target.appendChild(new_div)

        }

    }

    public async update(options: VisualUpdateOptions) {

        // Clear the event listener from the input element 
        const inputNode: Node = await this.clearDomNode(this.getInputDivElement())

        const metadataColumn: DataViewMetadataColumn = options.dataViews[0].table.columns[0];

        const target: IFilterColumnTarget = {
            table: metadataColumn.queryName.substr(0, metadataColumn.queryName.indexOf('.')), // table
            column: metadataColumn.displayName // col1
        };

        const searchDebounced = debounce(
            () => this.inputNodeSearchCallback.bind(this)(inputNode as HTMLInputElement, target),
            this.searchDebounce,
        )

        inputNode.addEventListener('input', searchDebounced)

    }

    private clearDomNode(domNode: Node): Node {

        const replacement_domNode: Node = domNode.cloneNode(false);
        domNode.parentNode.replaceChild(replacement_domNode, domNode);
        return replacement_domNode

    }

    private inputNodeSearchCallback(inputNode: HTMLInputElement, target: IFilterColumnTarget): void {

        const currentInput: string = (inputNode).value

        // Handle Empty input string and clear the selections
        if (!currentInput) {
            this.host.applyJsonFilter(null, "general", "filter", FilterAction.merge);
            return
        }

        const conditions: IAdvancedFilterCondition[] = [];

        conditions.push({
            operator: "StartsWith",
            value: currentInput
        });

        const advancedFilterLogicalOperators: AdvancedFilterLogicalOperators = "And";

        const filter: AdvancedFilter = new AdvancedFilter(target, advancedFilterLogicalOperators, conditions);

        this.host.applyJsonFilter(filter, "general", "filter", FilterAction.merge);

    }

    /**
     * Gets the input search box 
     */
    private getInputDivElement(): Node {
        return this.target.getElementsByClassName("InputDiv")[0].firstChild
    }

    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property. 
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}