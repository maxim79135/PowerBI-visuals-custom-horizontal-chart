/*
 *  Power BI Visualizations
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

import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;

export class BarShape {
  public color: string = "#01b8aa";
  public showAll: boolean = false;
  public minHeight: number = 30;
  public barPadding: number = 15;
}

export class XAxis {
  public textSize: number = 10;
  public fontFamily: string = "segoe";
  public categoryColor: string = "#333333";
  public rangeColor: string = "#333333";
  public isItalic: boolean = false;
  public isBold: boolean = false;
  public decimalPlaces: number = 0;
  // public wordWrap_: boolean = false;
}

export class YAxis {
  public textPosition: string = "insideBar";
  public textSize: number = 10;
  public fontFamily: string = "segoe";
  public fontColor: string = "#333333";
  public isItalic: boolean = false;
  public isBold: boolean = false;
  // public wordWrap_: boolean = false;
  public paddingLeft: number = 5;
  public displayUnit: number = 0;
  public decimalPlaces: number = 0;
}

export class BarSettings extends DataViewObjectsParser {
  public barShape: BarShape = new BarShape();
  public xAxis: XAxis = new XAxis();
  public yAxis: YAxis = new YAxis();
}
