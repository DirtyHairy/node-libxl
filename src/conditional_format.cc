/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2025 Christian Speckner <cnspeckn@googlemail.com>
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

#include "conditional_format.h"

#include "argument_helper.h"
#include "assert.h"
#include "font.h"
#include "util.h"

using namespace v8;

namespace node_libxl {

    ConditionalFormat::ConditionalFormat(libxl::ConditionalFormat* conditionalFormat,
                                         Local<Value> book)
        : Wrapper<libxl::ConditionalFormat, ConditionalFormat>(conditionalFormat),
          BookHolder(book) {}

    Local<Object> ConditionalFormat::NewInstance(libxl::ConditionalFormat* libxlConditionalFormat,
                                                 Local<Value> book) {
        Nan::EscapableHandleScope scope;

        ConditionalFormat* conditionalFormat = new ConditionalFormat(libxlConditionalFormat, book);
        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();
        conditionalFormat->Wrap(that);

        return scope.Escape(that);
    }

    // Method implementations

    NAN_METHOD(ConditionalFormat::Font) {
        Nan::HandleScope scope;

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        libxl::Font* font = that->GetWrapped()->font();
        if (!font) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Font::NewInstance(font, that->GetBookHandle()));
    }

    NAN_METHOD(ConditionalFormat::NumFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->numFormat()));
    }

    NAN_METHOD(ConditionalFormat::SetNumFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int format = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setNumFormat(static_cast<libxl::NumFormat>(format));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::CustomNumFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* format = that->GetWrapped()->customNumFormat();
        if (!format) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(format).ToLocalChecked());
    }

    NAN_METHOD(ConditionalFormat::SetCustomNumFormat) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(format, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCustomNumFormat(*format);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::SetBorder) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int style = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorder(static_cast<libxl::BorderStyle>(style));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::SetBorderColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderLeft()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderLeft) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int style = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderLeft(static_cast<libxl::BorderStyle>(style));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderRight()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderRight) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int style = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderRight(static_cast<libxl::BorderStyle>(style));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderTop) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderTop()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderTop) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int style = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderTop(static_cast<libxl::BorderStyle>(style));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderBottom()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderBottom) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int style = arguments.GetInt(0, libxl::BORDERSTYLE_THIN);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderBottom(static_cast<libxl::BorderStyle>(style));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderLeftColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderLeftColor()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderLeftColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderLeftColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderRightColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderRightColor()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderRightColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderRightColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderTopColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderTopColor()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderTopColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderTopColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::BorderBottomColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->borderBottomColor()));
    }

    NAN_METHOD(ConditionalFormat::SetBorderBottomColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBorderBottomColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::FillPattern) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->fillPattern()));
    }

    NAN_METHOD(ConditionalFormat::SetFillPattern) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int pattern = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFillPattern(static_cast<libxl::FillPattern>(pattern));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::PatternForegroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->patternForegroundColor()));
    }

    NAN_METHOD(ConditionalFormat::SetPatternForegroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setPatternForegroundColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormat::PatternBackgroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->patternBackgroundColor()));
    }

    NAN_METHOD(ConditionalFormat::SetPatternBackgroundColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormat* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setPatternBackgroundColor(static_cast<libxl::Color>(color));
        info.GetReturnValue().Set(info.This());
    }

    void ConditionalFormat::Initialize(Local<Object> exports) {
        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("ConditionalFormat").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<ConditionalFormat>(t);

        Nan::SetPrototypeMethod(t, "font", Font);
        Nan::SetPrototypeMethod(t, "numFormat", NumFormat);
        Nan::SetPrototypeMethod(t, "setNumFormat", SetNumFormat);
        Nan::SetPrototypeMethod(t, "customNumFormat", CustomNumFormat);
        Nan::SetPrototypeMethod(t, "setCustomNumFormat", SetCustomNumFormat);
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
        Nan::SetPrototypeMethod(t, "fillPattern", FillPattern);
        Nan::SetPrototypeMethod(t, "setFillPattern", SetFillPattern);
        Nan::SetPrototypeMethod(t, "patternForegroundColor", PatternForegroundColor);
        Nan::SetPrototypeMethod(t, "setPatternForegroundColor", SetPatternForegroundColor);
        Nan::SetPrototypeMethod(t, "patternBackgroundColor", PatternBackgroundColor);
        Nan::SetPrototypeMethod(t, "setPatternBackgroundColor", SetPatternBackgroundColor);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("ConditionalFormat").ToLocalChecked(),
                 Nan::New(constructor));
    }
}  // namespace node_libxl
