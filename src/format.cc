/**
 * The MIT License (MIT)
 * 
 * Copyright (c) 2013 Christian Speckner <cnspeckn@googlemail.com>
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

#include "format.h"

#include "assert.h"
#include "util.h"
#include "argument_helper.h"
#include "font.h"

using namespace v8;

namespace node_libxl {


// Lifecycle

Format::Format(libxl::Format* format, Handle<Value> book) :
    Wrapper<libxl::Format>(format),
    BookWrapper(book)
{}


Handle<Object> Format::NewInstance(
    libxl::Format* libxlFormat,
    Handle<Value> book)
{
    NanEscapableScope();

    Format* format = new Format(libxlFormat, book);

    Local<Object> that = 
        NanNew(util::CallStubConstructor(NanNew(constructor)).As<Object>());

    format->Wrap(that);

    return NanEscapeScope(that);
}


// Wrappers


NAN_METHOD(Format::Font) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Font* font = that->GetWrapped()->font();
    if (!font) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(Font::NewInstance(font, that->GetBookHandle()));
}


NAN_METHOD(Format::SetFont) {
    NanScope();

    ArgumentHelper arguments(args);

    node_libxl::Font* font = arguments.GetWrapped<node_libxl::Font>(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setFont(font->GetWrapped())) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetNumFormat) {
    NanScope();

    ArgumentHelper arguments(args);
    
    int format = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setNumFormat(format);

    NanReturnValue(args.This());
}


NAN_METHOD(Format::NumFormat) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->numFormat()));
}


NAN_METHOD(Format::AlignH) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->alignH()));
}


NAN_METHOD(Format::SetAlignH) {
    NanScope();

    ArgumentHelper arguments(args);

    int align = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setAlignH(static_cast<libxl::AlignH>(align));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::GetWrap) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->wrap()));
}


NAN_METHOD(Format::SetWrap) {
    NanScope();

    ArgumentHelper arguments(args);
    bool wrap = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setWrap(wrap);

    NanReturnValue(args.This());
}


NAN_METHOD(Format::Rotation) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->rotation()));
}


NAN_METHOD(Format::SetRotation) {
    NanScope();

    ArgumentHelper arguments(args);

    int rotation = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setRotation(rotation)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Format::Indent) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->indent()));
}


NAN_METHOD(Format::SetIndent) {
    NanScope();

    ArgumentHelper arguments(args);

    int indent = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setIndent(indent);

    NanReturnValue(args.This());
}

NAN_METHOD(Format::ShrinkToFit) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->shrinkToFit()));
}


NAN_METHOD(Format::SetShrinkToFit) {
    NanScope();

    ArgumentHelper arguments(args);

    bool shrink = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setShrinkToFit(shrink);

    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetBorder) {
    NanScope();

    ArgumentHelper arguments(args);

    int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorder(static_cast<libxl::BorderStyle>(border));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetBorderColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderColor(static_cast<libxl::Color>(color));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderLeft) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderLeft()));
}


NAN_METHOD(Format::SetBorderLeft) {
    NanScope();

    ArgumentHelper arguments(args);

    int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderLeft(static_cast<libxl::BorderStyle>(border));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderRight) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderRight()));
}


NAN_METHOD(Format::SetBorderRight) {
    NanScope();

    ArgumentHelper arguments(args);

    int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderRight(static_cast<libxl::BorderStyle>(border));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderTop) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderTop()));
}


NAN_METHOD(Format::SetBorderTop) {
    NanScope();

    ArgumentHelper arguments(args);

    int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderTop(static_cast<libxl::BorderStyle>(border));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderBottom) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderBottom()));
}


NAN_METHOD(Format::SetBorderBottom) {
    NanScope();

    ArgumentHelper arguments(args);

    int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderBottom(static_cast<libxl::BorderStyle>(border));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderLeftColor) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderLeftColor()));
}


NAN_METHOD(Format::SetBorderLeftColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderLeftColor(static_cast<libxl::Color>(color));
    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderRightColor) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderRightColor()));
}


NAN_METHOD(Format::SetBorderRightColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderRightColor(static_cast<libxl::Color>(color));
    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderTopColor) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderTopColor()));
}


NAN_METHOD(Format::SetBorderTopColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderTopColor(static_cast<libxl::Color>(color));
    NanReturnValue(args.This());
}


NAN_METHOD(Format::BorderBottomColor) {
    NanScope();

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->borderBottomColor()));
}


NAN_METHOD(Format::SetBorderBottomColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBorderBottomColor(static_cast<libxl::Color>(color));
    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetFillPattern) {
    NanScope();

    ArgumentHelper arguments(args);

    int pattern = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setFillPattern(static_cast<libxl::FillPattern>(pattern));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetPatternBackgroundColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPatternBackgroundColor(static_cast<libxl::Color>(color));

    NanReturnValue(args.This());
}


NAN_METHOD(Format::SetPatternForegroundColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Format* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setPatternForegroundColor(static_cast<libxl::Color>(color));

    NanReturnValue(args.This());
}


// Init


void Format::Initialize(Handle<Object> exports) {
    using namespace libxl;
    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(util::StubConstructor);
    t->SetClassName(NanNew<String>("Format"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    BookWrapper::Initialize<Format>(t);

    NODE_SET_PROTOTYPE_METHOD(t, "font", Font);
    NODE_SET_PROTOTYPE_METHOD(t, "setFont", SetFont);
    NODE_SET_PROTOTYPE_METHOD(t, "setNumFormat", SetNumFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "numFormat", NumFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "alignH", AlignH);
    NODE_SET_PROTOTYPE_METHOD(t, "setAlignH", SetAlignH);
    NODE_SET_PROTOTYPE_METHOD(t, "wrap", GetWrap);
    NODE_SET_PROTOTYPE_METHOD(t, "setWrap", SetWrap);
    NODE_SET_PROTOTYPE_METHOD(t, "rotation", Rotation);
    NODE_SET_PROTOTYPE_METHOD(t, "setRotation", SetRotation);
    NODE_SET_PROTOTYPE_METHOD(t, "indent", Indent);
    NODE_SET_PROTOTYPE_METHOD(t, "setIndent", SetIndent);
    NODE_SET_PROTOTYPE_METHOD(t, "shrinkToFit", ShrinkToFit);
    NODE_SET_PROTOTYPE_METHOD(t, "setShrinkToFit", SetShrinkToFit);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorder", SetBorder);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderColor", SetBorderColor);
    NODE_SET_PROTOTYPE_METHOD(t, "borderLeft", BorderLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderLeft", SetBorderLeft);
    NODE_SET_PROTOTYPE_METHOD(t, "borderRight", BorderRight);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderRight", SetBorderRight);
    NODE_SET_PROTOTYPE_METHOD(t, "borderTop", BorderTop);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderTop", SetBorderTop);
    NODE_SET_PROTOTYPE_METHOD(t, "borderBottom", BorderBottom);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderBottom", SetBorderBottom);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderLeftColor", SetBorderLeftColor);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderRightColor", SetBorderRightColor);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderTopColor", SetBorderTopColor);
    NODE_SET_PROTOTYPE_METHOD(t, "setBorderBottomColor", SetBorderBottomColor);
    NODE_SET_PROTOTYPE_METHOD(t, "setFillPattern", SetFillPattern);
    NODE_SET_PROTOTYPE_METHOD(t, "setPatternBackgroundColor",
        SetPatternBackgroundColor);
    NODE_SET_PROTOTYPE_METHOD(t, "setPatternForegroundColor",
        SetPatternForegroundColor);

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());

    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_GENERAL);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_D2);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_SEP);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_SEP_D2);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CURRENCY_NEGBRA);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CURRENCY_NEGBRARED);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CURRENCY_D2_NEGBRA);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CURRENCY_D2_NEGBRARED);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_PERCENT);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_PERCENT_D2);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_SCIENTIFIC_D2);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_FRACTION_ONEDIG);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_FRACTION_TWODIG);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_DATE);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_D_MON_YY);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_D_MON);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_MON_YY);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_HMM_AM);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_HMMSS_AM);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_HMM);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_HMMSS);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_MDYYYY_HMM);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_SEP_NEGBRA);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_SEP_NEGBRARED);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_D2_SEP_NEGBRA);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_NUMBER_D2_SEP_NEGBRARED);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_ACCOUNT);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_ACCOUNTCUR);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_ACCOUNT_D2);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_ACCOUNT_D2_CUR);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_MMSS);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_H0MMSS);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_MMSS0);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_CUSTOM_000P0E_PLUS0);
    NODE_DEFINE_CONSTANT(exports, NUMFORMAT_TEXT);

    NODE_DEFINE_CONSTANT(exports, ALIGNH_GENERAL);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_LEFT);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_CENTER);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_RIGHT);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_FILL);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_JUSTIFY);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_MERGE);
    NODE_DEFINE_CONSTANT(exports, ALIGNH_DISTRIBUTED);

    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_NONE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_SOLID);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_GRAY50);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_GRAY75);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_GRAY25);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_HORSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_VERSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_REVDIAGSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_DIAGSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_DIAGCROSSHATCH);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THICKDIAGCROSSHATCH);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINHORSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINVERSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINREVDIAGSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINDIAGSTRIPE);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINHORCROSSHATCH);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_THINDIAGCROSSHATCH);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_GRAY12P5);
    NODE_DEFINE_CONSTANT(exports, FILLPATTERN_GRAY6P25);

    NODE_DEFINE_CONSTANT(exports, COLOR_BLACK);
    NODE_DEFINE_CONSTANT(exports, COLOR_WHITE);
    NODE_DEFINE_CONSTANT(exports, COLOR_RED);
    NODE_DEFINE_CONSTANT(exports, COLOR_BRIGHTGREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_BLUE);
    NODE_DEFINE_CONSTANT(exports, COLOR_YELLOW);
    NODE_DEFINE_CONSTANT(exports, COLOR_PINK);
    NODE_DEFINE_CONSTANT(exports, COLOR_TURQUOISE);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKRED);
    NODE_DEFINE_CONSTANT(exports, COLOR_GREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKBLUE);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKYELLOW);
    NODE_DEFINE_CONSTANT(exports, COLOR_VIOLET);
    NODE_DEFINE_CONSTANT(exports, COLOR_TEAL);
    NODE_DEFINE_CONSTANT(exports, COLOR_GRAY25);
    NODE_DEFINE_CONSTANT(exports, COLOR_GRAY50);
    NODE_DEFINE_CONSTANT(exports, COLOR_PERIWINKLE_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_PLUM_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_IVORY_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTTURQUOISE_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKPURPLE_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_CORAL_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_OCEANBLUE_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_ICEBLUE_CF);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKBLUE_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_PINK_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_YELLOW_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_TURQUOISE_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_VIOLET_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKRED_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_TEAL_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_BLUE_CL);
    NODE_DEFINE_CONSTANT(exports, COLOR_SKYBLUE);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTTURQUOISE);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTGREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTYELLOW);
    NODE_DEFINE_CONSTANT(exports, COLOR_PALEBLUE);
    NODE_DEFINE_CONSTANT(exports, COLOR_ROSE);
    NODE_DEFINE_CONSTANT(exports, COLOR_LAVENDER);
    NODE_DEFINE_CONSTANT(exports, COLOR_TAN);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTBLUE);
    NODE_DEFINE_CONSTANT(exports, COLOR_AQUA);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIME);
    NODE_DEFINE_CONSTANT(exports, COLOR_GOLD);
    NODE_DEFINE_CONSTANT(exports, COLOR_LIGHTORANGE);
    NODE_DEFINE_CONSTANT(exports, COLOR_ORANGE);
    NODE_DEFINE_CONSTANT(exports, COLOR_BLUEGRAY);
    NODE_DEFINE_CONSTANT(exports, COLOR_GRAY40);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKTEAL);
    NODE_DEFINE_CONSTANT(exports, COLOR_SEAGREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_DARKGREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_OLIVEGREEN);
    NODE_DEFINE_CONSTANT(exports, COLOR_BROWN);
    NODE_DEFINE_CONSTANT(exports, COLOR_PLUM);
    NODE_DEFINE_CONSTANT(exports, COLOR_INDIGO);
    NODE_DEFINE_CONSTANT(exports, COLOR_GRAY80);
    NODE_DEFINE_CONSTANT(exports, COLOR_DEFAULT_FOREGROUND);
    NODE_DEFINE_CONSTANT(exports, COLOR_DEFAULT_BACKGROUND);
    NODE_DEFINE_CONSTANT(exports, COLOR_TOOLTIP);
    NODE_DEFINE_CONSTANT(exports, COLOR_AUTO);

    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_NONE);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_THIN);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_MEDIUM);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_DASHED);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_DOTTED);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_THICK);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_DOUBLE);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_HAIR);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_MEDIUMDASHED);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_DASHDOT);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_MEDIUMDASHDOT);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_DASHDOTDOT);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_MEDIUMDASHDOTDOT);
    NODE_DEFINE_CONSTANT(exports, BORDERSTYLE_SLANTDASHDOT);
}


}
