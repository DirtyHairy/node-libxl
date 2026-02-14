import { ConditionalFormat } from './conditional_format';

export class ConditionalFormatting {
    addRange(rowFirst: number, rowLast: number, colFirst: number, colLast: number): ConditionalFormatting;
    addRule(
        type: number,
        conditionalFormat: ConditionalFormat,
        value: string,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    addTopRule(
        conditionalFormat: ConditionalFormat,
        value: number,
        bottom?: boolean,
        percent?: boolean,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    addOpNumRule(
        op: number,
        conditionalFormat: ConditionalFormat,
        value1: number,
        value2?: number,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    addOpStrRule(
        op: number,
        conditionalFormat: ConditionalFormat,
        value1: string,
        value2?: string,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    addAboveAverageRule(
        conditionalFormat: ConditionalFormat,
        aboveAverage?: boolean,
        equalAverage?: boolean,
        stdDev?: number,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    addTimePeriodRule(
        conditionalFormat: ConditionalFormat,
        timePeriod: number,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    add2ColorScaleRule(
        minColor: number,
        maxColor: number,
        minType?: number,
        minValue?: number,
        maxType?: number,
        maxValue?: number,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    add2ColorScaleFormulaRule(
        minColor: number,
        maxColor: number,
        minType: number,
        minValue: string,
        maxType: number,
        maxValue: string,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    add3ColorScaleRule(
        minColor: number,
        midColor: number,
        maxColor: number,
        minType?: number,
        minValue?: number,
        midType?: number,
        midValue?: number,
        maxType?: number,
        maxValue?: number,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
    add3ColorScaleFormulaRule(
        minColor: number,
        midColor: number,
        maxColor: number,
        minType: number,
        minValue: string,
        midType: number,
        midValue: string,
        maxType: number,
        maxValue: string,
        stopIfTrue?: boolean,
    ): ConditionalFormatting;
}
