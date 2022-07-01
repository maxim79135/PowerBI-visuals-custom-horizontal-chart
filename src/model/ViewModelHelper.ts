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

import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import ValueTypeDescriptor = powerbi.ValueTypeDescriptor;

function parseSettings(dataView: DataView): BarSettings {
  return <BarSettings>BarSettings.parse(dataView);
}

// tslint:disable-next-line: export-name
export function visualTransform(
  options: VisualUpdateOptions,
  host: IVisualHost
): IBarChartViewModel {
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
      dataPoints: dataPoints,
      tooltipValues: [],
    };

  let dataCategorical = dataViews[0].categorical;
  let category = dataCategorical.categories
    ? dataCategorical.categories[dataCategorical.categories.length - 1]
    : null;
  let categories = category ? category.values : [""];
  let dataMax: number;

  for (let i = 0; i < categories.length; i++) {
    let dataPoint = <IDataPoint>{};

    for (let ii = 0; ii < dataCategorical.values!.length; ii++) {
      let dataValue = dataCategorical.values![ii];
      let value: any = dataValue.values[i];
      let valueType = dataValue.source.type;

      if (!dataValue.source.roles) break;
      if (!valueType) valueType = {};

      if (dataValue.source.roles["measure"]) {
        if (categories[i]) {
          dataPoint.category = <string>categories[i];
        } else {
          if (category) dataPoint.category = "";
          else dataPoint.category = dataValue.source.displayName;
        }

        dataPoint.id = i;
        dataPoint.value = valueType.numeric || valueType.integer ? value : null;

        dataMax = <number>dataValue.maxLocal;
      }
    }

    dataPoints.push(dataPoint);
  }

  return {
    settings: settings,
    dataMax: dataMax,
    dataPoints: dataPoints,
    tooltipValues: [],
  };
}
