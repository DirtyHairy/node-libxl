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

using namespace v8;

namespace node_libxl {


// Lifecycle

Format::Format(libxl::Format* format, Handle<Value> book) :
    Wrapper<libxl::Format>(format),
    bookHandle(Persistent<Value>::New(book))
{}


Format::~Format() {
    bookHandle.Dispose();
}


Handle<Object> Format::NewInstance(
    libxl::Format* libxlFormat,
    Handle<Value> book)
{
    HandleScope scope;

    Format* format = new Format(libxlFormat, book);

    Handle<Object> that = util::CallStubConstructor(constructor).As<Object>();

    format->Wrap(that);

    return scope.Close(that);
}


// Wrappers


Handle<Value> Format::SetNumFormat(const Arguments& arguments) {
    HandleScope scope;

    ArgumentHelper args(arguments);
    
    int32_t format = args.GetInt(0);
    ASSERT_ARGUMENTS(args);

    Format* that = Unwrap(arguments.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setNumFormat(format);

    return scope.Close(arguments.This());
}


// Init


void Format::Initialize(Handle<Object> exports) {
    HandleScope scope;

    using namespace libxl;

    Local<FunctionTemplate> t = FunctionTemplate::New(util::StubConstructor);
    t->SetClassName(String::NewSymbol("Format"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    NODE_SET_PROTOTYPE_METHOD(t, "setNumFormat", SetNumFormat);

    t->ReadOnlyPrototype();
    constructor = Persistent<Function>::New(t->GetFunction());

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
}


}
