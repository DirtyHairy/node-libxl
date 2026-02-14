export class Font {
    size(): number;
    setSize(size: number): Font;
    italic(): boolean;
    setItalic(italic?: boolean): Font;
    strikeOut(): boolean;
    setStrikeOut(strikeOut?: boolean): Font;
    color(): number;
    setColor(color: number): Font;
    bold(): boolean;
    setBold(bold?: boolean): Font;
    script(): number;
    setScript(script: number): Font;
    underline(): number;
    setUnderline(underline: number): Font;
    name(): string;
    setName(name: string): Font;
}
