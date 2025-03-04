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

#include "argument_helper.h"
#include "assert.h"
#include "font.h"
#include "util.h"

using namespace v8;

namespace node_libxl {

    // Lifecycle

    Format::Format(libxl::Format* format, Local<Value> book)
        : Wrapper<libxl::Format, Format>(format), BookHolder(book) {}

    Local<Object> Format::NewInstance(libxl::Format* libxlFormat, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        Format* format = new Format(libxlFormat, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        format->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(Format::Font) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::Font* font = that->GetWrapped()->font();
        if (!font) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Font::NewInstance(font, that->GetBookHandle()));
    }

    NAN_METHOD(Format::SetFont) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        node_libxl::Font* font = arguments.GetWrapped<node_libxl::Font>(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->setFont(font->GetWrapped())) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::SetNumFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int format = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setNumFormat(format);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::NumFormat) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->numFormat()));
    }

    NAN_METHOD(Format::AlignH) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->alignH()));
    }

    NAN_METHOD(Format::SetAlignH) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int align = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setAlignH(static_cast<libxl::AlignH>(align));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::AlignV) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->alignV()));
    }

    NAN_METHOD(Format::SetAlignV) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int align = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setAlignV(static_cast<libxl::AlignV>(align));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::GetWrap) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->wrap()));
    }

    NAN_METHOD(Format::SetWrap) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        bool wrap = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setWrap(wrap);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::Rotation) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->rotation()));
    }

    NAN_METHOD(Format::SetRotation) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rotation = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->setRotation(rotation)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::Indent) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->indent()));
    }

    NAN_METHOD(Format::SetIndent) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int indent = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setIndent(indent);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::ShrinkToFit) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->shrinkToFit()));
    }

    NAN_METHOD(Format::SetShrinkToFit) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool shrink = arguments.GetBoolean(0, true);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setShrinkToFit(shrink);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::SetBorder) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorder(static_cast<libxl::BorderStyle>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::SetBorderColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderColor(static_cast<libxl::Color>(color));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderLeft) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderLeft()));
    }

    NAN_METHOD(Format::SetBorderLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderLeft(static_cast<libxl::BorderStyle>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderRight) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderRight()));
    }

    NAN_METHOD(Format::SetBorderRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderRight(static_cast<libxl::BorderStyle>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderTop) {
        Nan::HandleScope scope;

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderTop()));
    }

    NAN_METHOD(Format::SetBorderTop) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderTop(static_cast<libxl::BorderStyle>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderBottom()));
    }

    NAN_METHOD(Format::SetBorderBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderBottom(static_cast<libxl::BorderStyle>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderLeftColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderLeftColor()));
    }

    NAN_METHOD(Format::SetBorderLeftColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderLeftColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderRightColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderRightColor()));
    }

    NAN_METHOD(Format::SetBorderRightColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderRightColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderTopColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderTopColor()));
    }

    NAN_METHOD(Format::SetBorderTopColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderTopColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderBottomColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderBottomColor()));
    }

    NAN_METHOD(Format::SetBorderBottomColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderBottomColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderDiagonal) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderDiagonal()));
    }

    NAN_METHOD(Format::SetBorderDiagonal) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int border = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderDiagonal(static_cast<libxl::BorderDiagonal>(border));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::BorderDiagonalColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->borderDiagonalColor()));
    }

    NAN_METHOD(Format::SetBorderDiagonalColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderDiagonalColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::FillPattern) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->fillPattern()));
    }

    NAN_METHOD(Format::SetFillPattern) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int pattern = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFillPattern(static_cast<libxl::FillPattern>(pattern));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::PatternBackgroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->patternBackgroundColor()));
    }

    NAN_METHOD(Format::SetPatternBackgroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setPatternBackgroundColor(static_cast<libxl::Color>(color));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::PatternForegroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->patternForegroundColor()));
    }

    NAN_METHOD(Format::SetPatternForegroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setPatternForegroundColor(static_cast<libxl::Color>(color));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::Locked) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->locked()));
    }

    NAN_METHOD(Format::SetLocked) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool locked = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setLocked(locked);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Format::Hidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->hidden()));
    }

    NAN_METHOD(Format::SetHidden) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool hidden = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Format* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setHidden(hidden);

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void Format::Initialize(Local<Object> exports) {
        using namespace libxl;
        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("Format").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<Format>(t);

        Nan::SetPrototypeMethod(t, "font", Font);
        Nan::SetPrototypeMethod(t, "setFont", SetFont);
        Nan::SetPrototypeMethod(t, "setNumFormat", SetNumFormat);
        Nan::SetPrototypeMethod(t, "numFormat", NumFormat);
        Nan::SetPrototypeMethod(t, "alignH", AlignH);
        Nan::SetPrototypeMethod(t, "setAlignH", SetAlignH);
        Nan::SetPrototypeMethod(t, "alignV", AlignV);
        Nan::SetPrototypeMethod(t, "setAlignV", SetAlignV);
        Nan::SetPrototypeMethod(t, "wrap", GetWrap);
        Nan::SetPrototypeMethod(t, "setWrap", SetWrap);
        Nan::SetPrototypeMethod(t, "rotation", Rotation);
        Nan::SetPrototypeMethod(t, "setRotation", SetRotation);
        Nan::SetPrototypeMethod(t, "indent", Indent);
        Nan::SetPrototypeMethod(t, "setIndent", SetIndent);
        Nan::SetPrototypeMethod(t, "shrinkToFit", ShrinkToFit);
        Nan::SetPrototypeMethod(t, "setShrinkToFit", SetShrinkToFit);
        Nan::SetPrototypeMethod(t, "setBorder", SetBorder);
        Nan::SetPrototypeMethod(t, "setBorderColor", SetBorderColor);
        Nan::SetPrototypeMethod(t, "borderLeft", BorderLeft);
        Nan::SetPrototypeMethod(t, "setBorderLeft", SetBorderLeft);
        Nan::SetPrototypeMethod(t, "borderRight", BorderRight);
        Nan::SetPrototypeMethod(t, "setBorderRight", SetBorderRight);
        Nan::SetPrototypeMethod(t, "borderTop", BorderTop);
        Nan::SetPrototypeMethod(t, "setBorderTop", SetBorderTop);
        Nan::SetPrototypeMethod(t, "borderBottom", BorderBottom);
        Nan::SetPrototypeMethod(t, "setBorderBottom", SetBorderBottom);
        Nan::SetPrototypeMethod(t, "borderLeftColor", BorderLeftColor);
        Nan::SetPrototypeMethod(t, "setBorderLeftColor", SetBorderLeftColor);
        Nan::SetPrototypeMethod(t, "borderRightColor", BorderRightColor);
        Nan::SetPrototypeMethod(t, "setBorderRightColor", SetBorderRightColor);
        Nan::SetPrototypeMethod(t, "borderTopColor", BorderTopColor);
        Nan::SetPrototypeMethod(t, "setBorderTopColor", SetBorderTopColor);
        Nan::SetPrototypeMethod(t, "borderBottomColor", BorderBottomColor);
        Nan::SetPrototypeMethod(t, "setBorderBottomColor", SetBorderBottomColor);
        Nan::SetPrototypeMethod(t, "borderDiagonal", BorderDiagonal);
        Nan::SetPrototypeMethod(t, "setBorderDiagonal", SetBorderDiagonal);
        Nan::SetPrototypeMethod(t, "borderDiagonalColor", BorderDiagonalColor);
        Nan::SetPrototypeMethod(t, "setBorderDiagonalColor", SetBorderDiagonalColor);
        Nan::SetPrototypeMethod(t, "fillPattern", FillPattern);
        Nan::SetPrototypeMethod(t, "setFillPattern", SetFillPattern);
        Nan::SetPrototypeMethod(t, "patternBackgroundColor", PatternBackgroundColor);
        Nan::SetPrototypeMethod(t, "setPatternBackgroundColor", SetPatternBackgroundColor);
        Nan::SetPrototypeMethod(t, "patternForegroundColor", PatternForegroundColor);
        Nan::SetPrototypeMethod(t, "setPatternForegroundColor", SetPatternForegroundColor);
        Nan::SetPrototypeMethod(t, "hidden", Hidden);
        Nan::SetPrototypeMethod(t, "setHidden", SetHidden);
        Nan::SetPrototypeMethod(t, "locked", Locked);
        Nan::SetPrototypeMethod(t, "setLocked", SetLocked);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("Format").ToLocalChecked(), Nan::New(constructor));
    }

}  // namespace node_libxl
