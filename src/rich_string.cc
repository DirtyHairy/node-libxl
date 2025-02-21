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

#include "rich_string.h"

#include "argument_helper.h"
#include "assert.h"
#include "font.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    RichString::RichString(libxl::RichString* richString, Local<Value> book)
        : Wrapper<libxl::RichString>(richString), BookWrapper(book) {}

    Local<Object> RichString::NewInstance(libxl::RichString* libxlRichString, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        RichString* richString = new RichString(libxlRichString, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        richString->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(RichString::AddFont) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        Font* initFont = arguments.GetWrapped<Font>(0, nullptr);
        ASSERT_ARGUMENTS(arguments);

        RichString* that = Unwrap(info.This());
        ASSERT_THIS(that);

        if (initFont) ASSERT_SAME_BOOK(initFont, that);

        libxl::Font* font =
            that->GetWrapped()->addFont(initFont ? initFont->GetWrapped() : nullptr);
        if (!font) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Font::NewInstance(font, that->GetBookHandle()));
    }

    NAN_METHOD(RichString::AddText) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(text, arguments.GetString(0));
        Font* font = arguments.GetWrapped<Font>(1, nullptr);

        ASSERT_ARGUMENTS(arguments);

        RichString* that = Unwrap(info.This());
        ASSERT_THIS(that);

        if (font) ASSERT_SAME_BOOK(font, that);

        that->GetWrapped()->addText(*text, font ? font->GetWrapped() : nullptr);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(RichString::GetText) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int index = arguments.GetInt(0);

        ASSERT_ARGUMENTS(arguments);

        RichString* that = Unwrap(info.This());
        ASSERT_THIS(that);

        libxl::Font* font = nullptr;

        const char* text = that->GetWrapped()->getText(index, &font);
        if (!text) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("text").ToLocalChecked(),
                 Nan::New<String>(text).ToLocalChecked());
        if (font) {
            Nan::Set(result, Nan::New<String>("font").ToLocalChecked(),
                     Font::NewInstance(font, that->GetBookHandle()));
        }

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(RichString::TextSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        RichString* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->textSize()));
    }

    // Init

    void RichString::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("RichString").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookWrapper::Initialize<RichString>(t);

        Nan::SetPrototypeMethod(t, "addFont", AddFont);
        Nan::SetPrototypeMethod(t, "addText", AddText);
        Nan::SetPrototypeMethod(t, "getText", GetText);
        Nan::SetPrototypeMethod(t, "textSize", TextSize);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("RichString").ToLocalChecked(), Nan::New(constructor));
    }
}  // namespace node_libxl
