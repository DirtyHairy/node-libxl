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

#include "font.h"

#include "argument_helper.h"
#include "assert.h"
#include "util.h"

using namespace v8;

namespace node_libxl {

    // Lifecycle

    Font::Font(libxl::Font* font, Local<Value> book)
        : Wrapper<libxl::Font, Font>(font), BookWrapper(book) {}

    Local<Object> Font::NewInstance(libxl::Font* libxlFont, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        Font* font = new Font(libxlFont, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        font->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(Font::SetSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int size = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setSize(size);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Size) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->size()));
    }

    NAN_METHOD(Font::Italic) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->italic()));
    }

    NAN_METHOD(Font::SetItalic) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool italic = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setItalic(italic);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::StrikeOut) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->strikeOut()));
    }

    NAN_METHOD(Font::SetStrikeOut) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool strikeOut = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setStrikeOut(strikeOut);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Color) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->color()));
    }

    NAN_METHOD(Font::SetColor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int color = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setColor(static_cast<libxl::Color>(color));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Bold) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->bold()));
    }

    NAN_METHOD(Font::SetBold) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool bold = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setBold(bold);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Script) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->script()));
    }

    NAN_METHOD(Font::SetScript) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int script = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setScript(static_cast<libxl::Script>(script));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Underline) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->underline()));
    }

    NAN_METHOD(Font::SetUnderline) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int underline = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setUnderline(static_cast<libxl::Underline>(underline));

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(Font::Name) {
        Nan::HandleScope scope;

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* name = that->GetWrapped()->name();
        if (!name) {
            return Nan::ThrowError("Unable to determine font name");
        }

        info.GetReturnValue().Set(Nan::New<String>(name).ToLocalChecked());
    }

    NAN_METHOD(Font::SetName) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(name, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        Font* that = Unwrap(info.This());
        ASSERT_THIS(that);

        if (!that->GetWrapped()->setName(*name)) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void Font::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("Font").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookWrapper::Initialize<Font>(t);

        Nan::SetPrototypeMethod(t, "setSize", SetSize);
        Nan::SetPrototypeMethod(t, "size", Size);
        Nan::SetPrototypeMethod(t, "italic", Italic);
        Nan::SetPrototypeMethod(t, "setItalic", SetItalic);
        Nan::SetPrototypeMethod(t, "strikeOut", StrikeOut);
        Nan::SetPrototypeMethod(t, "setStrikeOut", SetStrikeOut);
        Nan::SetPrototypeMethod(t, "color", Color);
        Nan::SetPrototypeMethod(t, "setColor", SetColor);
        Nan::SetPrototypeMethod(t, "bold", Bold);
        Nan::SetPrototypeMethod(t, "setBold", SetBold);
        Nan::SetPrototypeMethod(t, "script", Script);
        Nan::SetPrototypeMethod(t, "setScript", SetScript);
        Nan::SetPrototypeMethod(t, "underline", Underline);
        Nan::SetPrototypeMethod(t, "setUnderline", SetUnderline);
        Nan::SetPrototypeMethod(t, "name", Name);
        Nan::SetPrototypeMethod(t, "setName", SetName);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("Font").ToLocalChecked(), Nan::New(constructor));

        NODE_DEFINE_CONSTANT(exports, UNDERLINE_NONE);
        NODE_DEFINE_CONSTANT(exports, UNDERLINE_SINGLE);
        NODE_DEFINE_CONSTANT(exports, UNDERLINE_DOUBLE);
        NODE_DEFINE_CONSTANT(exports, UNDERLINE_SINGLEACC);
        NODE_DEFINE_CONSTANT(exports, UNDERLINE_DOUBLEACC);
        NODE_DEFINE_CONSTANT(exports, SCRIPT_NORMAL);
        NODE_DEFINE_CONSTANT(exports, SCRIPT_SUPER);
        NODE_DEFINE_CONSTANT(exports, SCRIPT_SUB);
    }

}  // namespace node_libxl
