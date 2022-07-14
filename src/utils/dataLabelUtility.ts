"use strict";

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import ValueTypeDescriptor = powerbi.ValueTypeDescriptor;

const localizedUnits = {
  "ru-RU_K": " тыс.",
  "ru-RU_M": " млн",
  "ru-RU_bn": " млрд",
  "ru-RU_T": " трлн",
};

function formatMeasure(data, properties) {
  const formatter = valueFormatter.create(properties);
  return formatter.format(data);
}

function localizeUnit(value: string, unit: string, culture: string): string {
  let localizedUnit = localizedUnits[culture + "_" + unit];
  if (localizedUnit) {
    return value.replace(unit, localizedUnit);
  } else return value;
}

function getValueType(valueType: ValueTypeDescriptor): string {
  let result: string = "";
  if (valueType.numeric || valueType.integer) result = "numeric";
  if (valueType.dateTime) result = "dateTime";
  return result;
}

// tslint:disable-next-line: export-name
export function prepareMeasureText(
  value: any,
  valueType: ValueTypeDescriptor,
  format: string,
  displayUnit: number,
  precision: number,
  addPlusforPositiveValue: boolean,
  suppressBlankAndNaN: boolean,
  blankAndNaNReplaceText: string,
  culture: string
): string {
  let valueFormatted: string = "";
  if (suppressBlankAndNaN) valueFormatted = blankAndNaNReplaceText;
  if (!(suppressBlankAndNaN && value == null)) {
    if (getValueType(valueType) == "numeric") {
      if (!(isNaN(value as number) && suppressBlankAndNaN)) {
        valueFormatted = formatMeasure(value as number, {
          format: format,
          value: displayUnit == 0 ? (value as number) : displayUnit,
          precision: precision,
          allowFormatBeautification: false,
          cultureSelector: culture,
        });
        if (culture != "en-US") {
          valueFormatted = localizeUnit(valueFormatted, "K", culture);
          valueFormatted = localizeUnit(valueFormatted, "M", culture);
          valueFormatted = localizeUnit(valueFormatted, "bn", culture);
          valueFormatted = localizeUnit(valueFormatted, "T", culture);
        }
        if (addPlusforPositiveValue && (value as number) > 0)
          valueFormatted = "+" + valueFormatted;
      }
      if (value == null && valueType["primitiveType"] == 3) {
        if (culture == "en-US") valueFormatted = "Infinity";
        else if (culture == "ru-RU") valueFormatted = "Бесконечность";
      }
    } else {
      valueFormatted = formatMeasure(
        getValueType(valueType) == "dateTime"
          ? new Date(value as string)
          : value,
        {
          format: format,
          cultureSelector: culture,
        }
      );
    }
  }

  if (valueFormatted == "(Blank)" && culture == "ru-RU")
    valueFormatted = "(Пусто)";

  return valueFormatted;
}
