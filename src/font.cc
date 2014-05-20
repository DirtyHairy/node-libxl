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


// Init


void Font::Initialize(Handle<Object> exports) {
    using namespace libxl;

    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(util::StubConstructor);
    t->SetClassName(NanSymbol("Font"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    BookWrapper::Initialize<Font>(t);

    NODE_SET_PROTOTYPE_METHOD(t, "setSize", SetSize);

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());
}

}
