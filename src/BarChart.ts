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
import TextProperties = interfaces.TextProperties;

import { dataViewWildcard } from "powerbi-visuals-utils-dataviewutils";
import { ColorHelper } from "powerbi-visuals-utils-colorutils";
import { BarSettings } from "./settings";
import { visualTransform } from "./model/ViewModelHelper";
import { IBarChartViewModel } from "./model/ViewModel";
import {
  textMeasurementService,
  interfaces,
} from "powerbi-visuals-utils-formattingutils";

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
    lineRangePadding: 2,
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

  private readonly outerPadding = 0.1;

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
    console.log(this.model);

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

    this.svg.selectAll("defs").remove();
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
    this.drawValueRangeShape();
    this.drawYAxis();
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
      .attr("fill-opacity", 1)
      .attr("selected", (d) => d.selected);

    bars.exit().remove();
    rects.exit().remove();
  }

  // tslint:disable-next-line: max-func-body-length
  public drawValueRangeShape() {
    const defs = this.svg.append("defs");

    let bars = this.barContainer.selectAll("g.bar").data(this.model.dataPoints);

    // drac background rect
    let backgroundRangeRect = bars
      .selectAll("rect.backgroundRangeRect")
      .data((d) => [d]);
    let mergeElement = backgroundRangeRect
      .enter()
      .append<SVGElement>("rect")
      .classed("backgroundRangeRect", true);
    backgroundRangeRect
      .merge(mergeElement)
      .attr("x", (d) => this.xScale(<number>d.minValue))
      .attr(
        "y",
        (d) => this.yScale(d.category) - BarChart.Config.lineRangePadding
      )
      .attr(
        "height",
        this.yScale.bandwidth() + 2 * BarChart.Config.lineRangePadding
      )
      .attr("width", (d) =>
        this.xScale(<number>d.maxValue - <number>d.minValue)
      )
      .style("fill", "#ffffff")
      .attr("fill-opacity", 0.5);

    // draw value range rect with pattern
    let valueRangesRect = bars
      .selectAll("rect.valueRangesRect")
      .data((d) => [d]);
    mergeElement = valueRangesRect
      .enter()
      .append<SVGElement>("rect")
      .classed("valueRangesRect", true);
    valueRangesRect
      .merge(mergeElement)
      .attr("x", (d) => this.xScale(<number>d.minValue))
      .attr(
        "y",
        (d) => this.yScale(d.category) - BarChart.Config.lineRangePadding
      )
      .attr(
        "height",
        this.yScale.bandwidth() + 2 * BarChart.Config.lineRangePadding
      )
      .attr("width", (d) =>
        this.xScale(<number>d.maxValue - <number>d.minValue)
      )
      .style("fill", (d) => {
        defs
          .append("pattern")
          .attr("id", d.category)
          .attr("width", "8")
          .attr("height", "8")
          .attr("patternUnits", "userSpaceOnUse")
          .attr("patternTransform", "rotate(-45)")
          .append("rect")
          .attr("width", "4")
          .attr("height", "8")
          .attr("transform", "translate(0,0)")
          .attr("fill", d.color);
        return "url(#" + d.category + ")";
      })
      .attr("fill-opacity", 0.5);

    if (this.model.dataPoints && this.model.dataPoints[0].minValue) {
      let minValueLine = bars.selectAll("line.minValueLine").data((d) => [d]);
      mergeElement = minValueLine
        .enter()
        .append<SVGElement>("line")
        .classed("minValueLine", true);
      minValueLine
        .merge(mergeElement)
        .attr("x1", (d) => this.xScale(<number>d.minValue))
        .attr(
          "y1",
          (d) => this.yScale(d.category) - BarChart.Config.lineRangePadding
        )
        .attr("x2", (d) => this.xScale(<number>d.minValue))
        .attr(
          "y2",
          (d) =>
            this.yScale(d.category) +
            this.yScale.bandwidth() +
            BarChart.Config.lineRangePadding
        )
        .style("stroke", (d) => d.color)
        .style("stroke-width", 4);
      minValueLine.exit().remove();
    } else bars.selectAll("line.minValueLine").remove();

    if (this.model.dataPoints && this.model.dataPoints[0].maxValue) {
      let maxValueLine = bars.selectAll("line.maxValueLine").data((d) => [d]);
      mergeElement = maxValueLine
        .enter()
        .append<SVGElement>("line")
        .classed("maxValueLine", true);
      maxValueLine
        .merge(mergeElement)
        .attr("x1", (d) => this.xScale(<number>d.maxValue))
        .attr(
          "y1",
          (d) => this.yScale(d.category) - BarChart.Config.lineRangePadding
        )
        .attr("x2", (d) => this.xScale(<number>d.maxValue))
        .attr(
          "y2",
          (d) =>
            this.yScale(d.category) +
            this.yScale.bandwidth() +
            BarChart.Config.lineRangePadding
        )
        .style("stroke", (d) => d.color)
        .style("stroke-width", 4);
      maxValueLine.exit().remove();
    } else bars.selectAll("line.maxValueLine").remove();

    defs.exit().remove();
    valueRangesRect.exit().remove();
  }

  public drawYAxis() {
    let settings = this.model.settings;
    let bars = this.barContainer.selectAll("g.bar").data(this.model.dataPoints);
    let yAxisText = bars.selectAll("text.yAxis-text").data((d) => [d]);
    let mergeElement = yAxisText
      .enter()
      .append<SVGElement>("text")
      .classed("yAxis-text", true);

    yAxisText
      .merge(mergeElement)
      .attr("x", settings.yAxis.paddingLeft)
      .attr("y", (d) => {
        let textProperties: TextProperties = {
          fontFamily: settings.yAxis.fontFamily,
          fontSize: settings.yAxis.textSize + "pt",
          text: d.formattedValue,
          fontWeight: settings.yAxis.isBold ? "bold" : "",
          fontStyle: settings.yAxis.isItalic ? "italic" : "",
        };
        return (
          this.yScale(d.category) +
          this.yScale.bandwidth() -
          textMeasurementService.measureSvgTextHeight(textProperties) / 4
        );
      })
      .attr("height", this.yScale.bandwidth())
      .attr("font-size", settings.yAxis.textSize)
      .attr("fill", settings.yAxis.fontColor)
      .text((d) => d.formattedValue)
      .each((d) => (d.width = this.xScale(<number>d.value)));

    yAxisText.exit().remove();
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

    if (
      options.objectName == "barShape" &&
      this.model.settings.barShape.showAll
    ) {
      this.enumerateCategories(instances, options.objectName);
    } else if (options.objectName == "yAxis") {
      this.enumerateYAxis(instances, options.objectName);
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

  private enumerateYAxis(
    instanceEnumeration: VisualObjectInstanceEnumeration,
    objectName: string
  ) {
    this.addAnInstanceToEnumeration(instanceEnumeration, {
      objectName,
      properties: {
        paddingLeft: this.model.settings.yAxis.paddingLeft,
      },
      selector: null,
      validValues: {
        paddingLeft: {
          numberRange: {
            min: 5,
            max: 25,
          },
        },
      },
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
