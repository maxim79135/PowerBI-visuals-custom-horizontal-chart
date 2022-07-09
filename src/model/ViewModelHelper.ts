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

import powerbi from "powerbi-visuals-api";
import { BarSettings } from "../settings";
import { IDataPoint, ITooltipValue, IBarChartViewModel } from "./ViewModel";
import { getValue } from "../utils/objectEnumerationUtility";
import { prepareMeasureText } from "../utils/dataLabelUtility";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import { ColorHelper } from "powerbi-visuals-utils-colorutils";

import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import ValueTypeDescriptor = powerbi.ValueTypeDescriptor;
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
import DataViewObjects = powerbi.DataViewObjects;
import IColorPalette = powerbi.extensibility.IColorPalette;
import PrimitiveValue = powerbi.PrimitiveValue;

function parseSettings(dataView: DataView): BarSettings {
  return <BarSettings>BarSettings.parse(dataView);
}

// tslint:disable-next-line: max-func-body-length disable-next-line: export-name
export function visualTransform(
  options: VisualUpdateOptions,
  host: IVisualHost
): IBarChartViewModel {
  const TeamsColorIdentifier: DataViewObjectPropertyIdentifier = {
    objectName: "teams",
    propertyName: "fill",
  };
  const colorPalette = host.colorPalette;

  let dataViews: DataView[] = options.dataViews;
  let dataPoints: IDataPoint[] = [];
  let settings: BarSettings = parseSettings(dataViews[0]);

  if (
    !dataViews ||
    !dataViews[0] ||
    !dataViews[0].categorical ||
    !dataViews[0].categorical.values
  )
    return {
      settings: settings,
      dataMax: 0,
      dataMin: 0,
      dataPoints: dataPoints,
    };

  let dataCategorical = dataViews[0].categorical;
  let category = dataCategorical.categories
    ? dataCategorical.categories[dataCategorical.categories.length - 1]
    : null;
  let categories = category ? category.values : [""];
  let dataMax: number = -1;
  let dataMin: number = Number.MAX_SAFE_INTEGER;

  for (let i = 0; i < categories.length; i++) {
    let dataPoint = <IDataPoint>{};
    dataPoint.tooltipValues = [];

    for (let ii = 0; ii < dataCategorical.values!.length; ii++) {
      let dataValue = dataCategorical.values![ii];
      let value: PrimitiveValue = dataValue.values[i];
      let maxLocal: PrimitiveValue = <number>dataValue.maxLocal;
      let minLocal: PrimitiveValue = <number>dataValue.minLocal;
      let valueType = dataValue.source.type;

      if (!dataValue.source.roles) break;
      if (!valueType) valueType = {};

      if (!maxLocal) maxLocal = <number>value;
      if (!minLocal) minLocal = <number>value;

      if (dataValue.source.roles["measure"]) {
        if (categories[i]) {
          dataPoint.category = <string>categories[i];
        } else {
          if (category) dataPoint.category = "";
          else dataPoint.category = dataValue.source.displayName;
        }

        dataPoint.id = i;
        dataPoint.value = valueType.numeric || valueType.integer ? value : null;
        dataPoint.formattedValue = prepareMeasureText(
          value,
          valueType,
          dataValue.objects
            ? <string>dataValue.objects[0]["general"]["formatString"]
            : valueFormatter.getFormatStringByColumn(dataValue.source),
          settings.yAxis.displayUnit,
          settings.yAxis.decimalPlaces,
          false,
          false,
          "",
          host.locale
        );
      }
      if (dataValue.source.roles["minMeasureValue"]) {
        dataPoint.minValue =
          valueType.numeric || valueType.integer ? value : null;
        dataPoint.minFormattedValue = prepareMeasureText(
          value,
          valueType,
          dataValue.objects
            ? <string>dataValue.objects[0]["general"]["formatString"]
            : valueFormatter.getFormatStringByColumn(dataValue.source),
          settings.yAxis.displayUnit,
          settings.categoryLabel.decimalPlaces,
          false,
          false,
          "",
          host.locale
        );
        dataPoint.displayNameMinValue = dataValue.source.displayName;
      }
      if (dataValue.source.roles["maxMeasureValue"]) {
        dataPoint.maxValue =
          valueType.numeric || valueType.integer ? value : null;
        dataPoint.maxFormattedValue = prepareMeasureText(
          value,
          valueType,
          dataValue.objects
            ? <string>dataValue.objects[0]["general"]["formatString"]
            : valueFormatter.getFormatStringByColumn(dataValue.source),
          settings.yAxis.displayUnit,
          settings.categoryLabel.decimalPlaces,
          false,
          false,
          "",
          host.locale
        );
        dataPoint.displayNameMaxValue = dataValue.source.displayName;
      }

      if (dataValue.source.roles["tooltip"]) {
        let tooltipValue: ITooltipValue = {
          displayName: dataValue.source.displayName,
          dataLabel: prepareMeasureText(
            value,
            valueType,
            dataValue.objects
              ? <string>dataValue.objects[0]["general"]["formatString"]
              : valueFormatter.getFormatStringByColumn(dataValue.source),
            1,
            0,
            false,
            false,
            "",
            "ru-RU"
          ),
        };
        dataPoint.tooltipValues.push(tooltipValue);
      }

      if (!dataValue.source.roles["tooltip"]) {
        if (maxLocal > dataMax) dataMax = maxLocal;
        if (minLocal < dataMin) dataMin = minLocal;
      }

      if (dataPoint.minValue > dataPoint.maxValue) {
        let tmp = dataPoint.minValue;
        dataPoint.minValue = dataPoint.maxValue;
        dataPoint.maxValue = tmp;

        tmp = dataPoint.minFormattedValue;
        dataPoint.minFormattedValue = dataPoint.maxFormattedValue;
        dataPoint.maxFormattedValue = tmp;
      }
      if (dataPoint.minValue && dataPoint.maxValue) {
        let minValueNew: string = dataPoint.minFormattedValue;
        if (
          dataPoint.minFormattedValue.endsWith("%") &&
          dataPoint.maxFormattedValue.endsWith("%")
        ) {
          minValueNew = dataPoint.minFormattedValue.split("%")[0];
        } else if (settings.yAxis.displayUnit > 1) {
          minValueNew = dataPoint.minFormattedValue.split(" ")[0];
        }
        dataPoint.rangeFormattedValue =
          minValueNew + "-" + dataPoint.maxFormattedValue;
      }
    }

    if (category && category.objects && category.objects[i]) {
      dataPoint.color = getValue(
        category.objects ? category.objects[i] : null,
        "barShape",
        "color",
        { solid: { color: "#333333" } }
      ).solid.color;
    } else {
      dataPoint.color = settings.barShape.color;
    }
    if (category) {
      dataPoint.selectionId = host
        .createSelectionIdBuilder()
        .withCategory(category, i)
        .createSelectionId();
    } else {
      dataPoint.selectionId = host
        .createSelectionIdBuilder()
        .withMeasure(dataCategorical.values[0].source.queryName)
        .createSelectionId();
    }

    dataPoints.push(dataPoint);
  }

  return {
    settings: settings,
    dataMax: dataMax,
    dataMin: dataMin,
    dataPoints: dataPoints,
  };
}
