import { Font } from "./font";

export class RichString {
    addFont(font?: Font): Font;
    addText(text: string, font?: Font): RichString;
    getText(index: number): { text: string; font: Font };
    textSize(): number;
}
