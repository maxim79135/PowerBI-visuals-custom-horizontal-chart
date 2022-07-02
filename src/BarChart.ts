/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
"use strict";

import "core-js/stable";
import "./../style/visual.less";

import { min } from "d3-array";
import { ScaleBand, scaleBand, ScaleLinear, scaleLinear } from "d3-scale";
import { BaseType, select, Selection } from "d3-selection";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualEnumerationInstanceKinds = powerbi.VisualEnumerationInstanceKinds;
import ISelectionId = powerbi.visuals.ISelectionId;
import { dataViewWildcard } from "powerbi-visuals-utils-dataviewutils";
import { ColorHelper } from "powerbi-visuals-utils-colorutils";

import { BarSettings } from "./settings";
import { visualTransform } from "./model/ViewModelHelper";
import { IBarChartViewModel } from "./model/ViewModel";

/**
 * An interface for reporting rendering events
 */
interface IVisualEventService {
  /**
   * Should be called just before the actual rendering was started.
   * Usually at the very start of the update method.
   *
   * @param options - the visual update options received as update parameter
   */
  renderingStarted(options: VisualUpdateOptions): void;

  /**
   * Shoudl be called immediately after finishing successfull rendering.
   *
   * @param options - the visual update options received as update parameter
   */
  renderingFinished(options: VisualUpdateOptions): void;

  /**
   * Called when rendering failed with optional reason string
   *
   * @param options - the visual update options received as update parameter
   * @param reason - the option failure reason string
   */
  renderingFailed(options: VisualUpdateOptions, reason?: string): void;
}

export class BarChart implements IVisual {
  // TEMP!
  private static Config = {
    barPadding: 0.15,
    fontScaleFactor: 3,
    maxHeightScale: 3,
    outerPaddingScale: 0.5,
    solidOpacity: 1,
    transparentOpacity: 0.5,
    xAxisFontMultiplier: 0.04,
    xScalePadding: 0.15,
    xScaledMin: 30,
  };

  private host: IVisualHost;
  private model: IBarChartViewModel;
  private events: IVisualEventService;

  private svg: Selection<SVGElement, {}, HTMLElement, any>;
  private divContainer: Selection<SVGElement, {}, HTMLElement, any>;
  private barContainer: Selection<SVGElement, {}, HTMLElement, any>;

  private width: number;
  private height: number;
  private yScale: ScaleBand<string>;
  private xScale: ScaleLinear<number, number, never>;

  private readonly outerPadding = -0.1;

  constructor(options: VisualConstructorOptions) {
    this.host = options.host;
    this.events = options.host.eventService;

    let svg = (this.svg = select(options.element)
      .append<SVGElement>("div")
      .classed("divContainer", true)
      .append<SVGElement>("svg")
      .classed("barChart", true));

    this.barContainer = svg
      .append<SVGElement>("g")
      .classed("barContainer", true);

    this.divContainer = select(".divContainer");
  }

  public update(options: VisualUpdateOptions) {
    this.model = visualTransform(options, this.host);
    this.width = options.viewport.width;
    this.height = options.viewport.height;

    this.events.renderingStarted(options);

    this.yScale = scaleBand()
      .domain(this.model.dataPoints.map((d) => d.category))
      .rangeRound([5, this.height])
      .padding(BarChart.Config.barPadding)
      .paddingOuter(this.outerPadding);
    // TEMP!
    let offset = this.width * 0.1;
    this.xScale = scaleLinear()
      .domain([0, this.model.dataMax])
      .range([0, this.width - offset - 40]); // subtracting 40 for padding between the bar and the label

    this.updateViewport(options);
    this.drawBarContainer();

    this.events.renderingFinished(options);
  }

  public updateViewport(options: VisualUpdateOptions) {
    let h = options.viewport.height + 5;
    let w = options.viewport.width;

    // update size canvas
    this.divContainer.attr(
      "style",
      "width:" + w + "px;height:" + h + "px;overflow-y:auto;overflow-x:hidden;"
    );
    this.svg.attr("width", this.width);
    this.svg.attr("height", this.height);

    // empty rect to take full width for clickable area for clearing selection
    let rectContainer = this.barContainer
      .selectAll("rect.rect-container")
      .data([0]);

    rectContainer
      .enter()
      .append<SVGElement>("rect")
      .classed("rect-container", true);

    rectContainer.attr("width", this.width);
    rectContainer.attr("height", this.height);
    rectContainer.attr("fill", "transparent");
  }

  public drawBarContainer() {
    let bars = this.barContainer.selectAll("g.bar").data(this.model.dataPoints);
    bars
      .enter()
      .append<SVGElement>("g")
      .classed("bar", true)
      .attr("x", BarChart.Config.xScalePadding) // .merge(bars)
      .attr("y", (d) => this.yScale(d.category))
      .attr("height", this.yScale.bandwidth())
      .attr("width", (d) => this.xScale(<number>d.value))

      .attr("selected", (d) => d.selected);
    this.drawBarShape();
  }

  public drawBarShape() {
    // create bar shape
    let bars = this.barContainer.selectAll("g.bar").data(this.model.dataPoints);
    let rects = bars.selectAll("rect.bar").data((d) => [d]);
    let mergeElement = rects
      .enter()
      .append<SVGElement>("rect")
      .classed("bar", true);

    rects
      .merge(mergeElement)
      .attr("x", BarChart.Config.xScalePadding)
      .attr("y", (d) => this.yScale(d.category))
      .attr("height", this.yScale.bandwidth())
      .attr("width", (d) => this.xScale(<number>d.value))
      .attr("fill", (d) => d.color)
      .attr("fill-opacity", 0.7)
      .attr("selected", (d) => d.selected);

    bars.exit().remove();
    rects.exit().remove();
  }

  /**
   * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
   * objects and properties you want to expose to the users in the property pane.
   *
   */
  public enumerateObjectInstances(
    options: EnumerateVisualObjectInstancesOptions
  ): VisualObjectInstanceEnumeration {
    const instances = BarSettings.enumerateObjectInstances(
      this.model.settings,
      options
    );
    console.log(instances);

    // const instances = [];

    if (
      options.objectName == "barShape" &&
      this.model.settings.barShape.showAll
    ) {
      this.enumerateCategories(instances, options.objectName);
    }

    return instances;
  }

  private enumerateCategories(
    instanceEnumeration: VisualObjectInstanceEnumeration,
    objectName: string
  ) {
    this.model.dataPoints.forEach((dataPoint) => {
      this.addAnInstanceToEnumeration(instanceEnumeration, {
        displayName: dataPoint.category,
        objectName: objectName,
        selector: ColorHelper.normalizeSelector(
          dataPoint.selectionId.getSelector(),
          false
        ),
        properties: {
          color: dataPoint.color,
        },
      });
    });
  }

  private addAnInstanceToEnumeration(
    instanceEnumeration: VisualObjectInstanceEnumeration,
    instance: VisualObjectInstance
  ): void {
    if (
      (<VisualObjectInstanceEnumerationObject>instanceEnumeration).instances
    ) {
      (<VisualObjectInstanceEnumerationObject>(
        instanceEnumeration
      )).instances.push(instance);
    } else {
      (<VisualObjectInstance[]>instanceEnumeration).push(instance);
    }
  }
}
