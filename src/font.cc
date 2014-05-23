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

#include "assert.h"
#include "util.h"
#include "argument_helper.h"

using namespace v8;

namespace node_libxl {


// Lifecycle


Font::Font(libxl::Font* font, Handle<Value> book) :
    Wrapper<libxl::Font>(font),
    BookWrapper(book)
{}


Handle<Object> Font::NewInstance(
    libxl::Font* libxlFont,
    Handle<Value> book)
{
    NanEscapableScope();

    Font* font = new Font(libxlFont, book);

    Local<Object> that = NanNew(util::CallStubConstructor(
        NanNew(constructor)).As<Object>());

    font->Wrap(that);

    return NanEscapeScope(that);
}


// Wrappers


NAN_METHOD(Font::SetSize) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t size = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setSize(size);

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Size) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->size()));
}


NAN_METHOD(Font::Italic) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->italic()));
}


NAN_METHOD(Font::SetItalic){
    NanScope();

    ArgumentHelper arguments(args);

    bool italic = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setItalic(italic);

    NanReturnValue(args.This());
}


NAN_METHOD(Font::StrikeOut) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->strikeOut()));
}


NAN_METHOD(Font::SetStrikeOut) {
    NanScope();

    ArgumentHelper arguments(args);

    bool strikeOut = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setStrikeOut(strikeOut);

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Color) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->color()));
}


NAN_METHOD(Font::SetColor) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t color = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setColor(static_cast<libxl::Color>(color));

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Bold) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->bold()));
}


NAN_METHOD(Font::SetBold) {
    NanScope();

    ArgumentHelper arguments(args);

    bool bold = arguments.GetBoolean(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setBold(bold);

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Script) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->script()));
}


NAN_METHOD(Font::SetScript) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t script = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setScript(static_cast<libxl::Script>(script));

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Underline) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->underline()));
}


NAN_METHOD(Font::SetUnderline) {
    NanScope();

    ArgumentHelper arguments(args);

    int32_t underline = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setUnderline(static_cast<libxl::Underline>(underline));

    NanReturnValue(args.This());
}


NAN_METHOD(Font::Name) {
    NanScope();

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    const char* name = that->GetWrapped()->name();
    if (!name) {
        return NanThrowError("Unable to determine font name");
    }

    NanReturnValue(NanNew<String>(name));
}


NAN_METHOD(Font::SetName) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Font* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->setName(*name)) {
        util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


// Init


void Font::Initialize(Handle<Object> exports) {
    using namespace libxl;

    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(util::StubConstructor);
    t->SetClassName(NanSymbol("Font"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    BookWrapper::Initialize<Font>(t);

    NODE_SET_PROTOTYPE_METHOD(t, "setSize", SetSize);
    NODE_SET_PROTOTYPE_METHOD(t, "size", Size);
    NODE_SET_PROTOTYPE_METHOD(t, "italic", Italic);
    NODE_SET_PROTOTYPE_METHOD(t, "setItalic", SetItalic);
    NODE_SET_PROTOTYPE_METHOD(t, "strikeOut", StrikeOut);
    NODE_SET_PROTOTYPE_METHOD(t, "setStrikeOut", SetStrikeOut);
    NODE_SET_PROTOTYPE_METHOD(t, "color", Color);
    NODE_SET_PROTOTYPE_METHOD(t, "setColor", SetColor);
    NODE_SET_PROTOTYPE_METHOD(t, "bold", Bold);
    NODE_SET_PROTOTYPE_METHOD(t, "setBold", SetBold);
    NODE_SET_PROTOTYPE_METHOD(t, "script", Script);
    NODE_SET_PROTOTYPE_METHOD(t, "setScript", SetScript);
    NODE_SET_PROTOTYPE_METHOD(t, "underline", Underline);
    NODE_SET_PROTOTYPE_METHOD(t, "setUnderline", SetUnderline);
    NODE_SET_PROTOTYPE_METHOD(t, "name", Name);
    NODE_SET_PROTOTYPE_METHOD(t, "setName", SetName);

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());

    NODE_DEFINE_CONSTANT(exports, UNDERLINE_NONE);
    NODE_DEFINE_CONSTANT(exports, UNDERLINE_SINGLE);
    NODE_DEFINE_CONSTANT(exports, UNDERLINE_DOUBLE);
    NODE_DEFINE_CONSTANT(exports, UNDERLINE_SINGLEACC);
    NODE_DEFINE_CONSTANT(exports, UNDERLINE_DOUBLEACC);
    NODE_DEFINE_CONSTANT(exports, SCRIPT_NORMAL);
    NODE_DEFINE_CONSTANT(exports, SCRIPT_SUPER);
    NODE_DEFINE_CONSTANT(exports, SCRIPT_SUB);    
}

}
