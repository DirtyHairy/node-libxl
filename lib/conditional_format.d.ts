import { Font } from "./font";

export class ConditionalFormat {
    font(): Font;
    numFormat(): number;
    setNumFormat(numFormat: number): ConditionalFormat;
    customNumFormat(): string;
    setCustomNumFormat(customNumFormat: string): ConditionalFormat;
    setBorderStyle(style: number): ConditionalFormat;
    borderLeft(): number;
    setBorderLeft(style: number): ConditionalFormat;
    borderRight(): number;
    setBorderRight(style: number): ConditionalFormat;
    borderTop(): number;
    setBorderTop(style: number): ConditionalFormat;
    borderBottom(): number;
    setBorderBottom(style: number): ConditionalFormat;
    setBorderColor(color: number): ConditionalFormat;
    borderLeftColor(): number;
    setBorderLeftColor(color: number): ConditionalFormat;
    borderRightColor(): number;
    setBorderRightColor(color: number): ConditionalFormat;
    borderTopColor(): number;
    setBorderTopColor(color: number): ConditionalFormat;
    borderBottomColor(): number;
    setBorderBottomColor(color: number): ConditionalFormat;
    fillPattern(): number;
    setFillPattern(pattern: number): ConditionalFormat;
    patternForegroundColor(): number;
    setPatternForegroundColor(color: number): ConditionalFormat;
    patternBackgroundColor(): number;
    setPatternBackgroundColor(color: number): ConditionalFormat;
}
