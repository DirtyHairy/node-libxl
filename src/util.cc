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

#include "util.h"

#include <libxl.h>
#include <nan.h>

#include "book.h"

using namespace v8;


namespace node_libxl {
namespace util {

Local<Value> ProxyConstructor(
    Handle<Function> constructor,
    Nan::NAN_METHOD_ARGS_TYPE arguments
) {
    Nan::EscapableHandleScope scope;

    uint8_t argc = arguments.Length();
    Handle<Value>* argv = new Handle<Value>[argc];

    for (uint8_t i = 0; i < argc; i++) {
        argv[i] = arguments[i];
    }

    Local<Value> newInstance = constructor->NewInstance(argc, argv);
    
    delete[] argv;
    return scope.Escape(newInstance);
}


NAN_METHOD(StubConstructor) {
    Nan::HandleScope scope;

    Handle<Value> sentry = info[0];

    if (!(  info.IsConstructCall() &&
            info.Length() == 1 &&
            sentry->IsExternal() &&
            sentry.As<External>()->Value() == NULL
    )) {
        return Nan::ThrowTypeError(
            "You are not supposed to call this constructor directly");
    }

    Nan::SetInternalFieldPointer(info.This(), 0, NULL);

    info.GetReturnValue().Set(info.This());
}


Local<Value> CallStubConstructor(Handle<Function> constructor) {
    Nan::EscapableHandleScope scope;

    Handle<Value> info[1] = {CSNanNewExternal(NULL)};

    return scope.Escape(constructor->NewInstance(1, info));
}


Book* GetBook(Book* book) {
    return book;
}


Book* GetBook(BookWrapper* bookWrapper) {
    return bookWrapper->GetBook();
}


libxl::Book* UnwrapBook(libxl::Book* book) {
    return book;
}


libxl::Book* UnwrapBook(v8::Handle<v8::Value> bookHandle) {
    Book* book = Book::Unwrap(bookHandle);

    return book ? book->GetWrapped() : NULL;
}


libxl::Book* UnwrapBook(Book* book) {
    return book ? book->GetWrapped() : NULL;
}


libxl::Book* UnwrapBook(BookWrapper* bookWrapper) {
    Book* book = bookWrapper->GetBook();
    return book ? book->GetWrapped() : NULL;
}


}
}
