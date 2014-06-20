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

#ifndef BINDINGS_UTIL
#define BINDINGS_UTIL

#include "common.h"
#include "book.h"
#include "book_wrapper.h"

namespace node_libxl {
namespace util {


v8::Handle<v8::Value> ProxyConstructor(
    v8::Handle<v8::Function> constructor,
    _NAN_METHOD_ARGS_TYPE arguments
);


template<typename T> _NAN_METHOD_RETURN_TYPE ThrowLibxlError(T book);


NAN_METHOD(StubConstructor);


v8::Handle<v8::Value> CallStubConstructor(v8::Handle<v8::Function> constructor);


bool AsyncPending(Book*);
bool AsyncPending(BookWrapper*);


libxl::Book* UnwrapBook(libxl::Book* book); 
libxl::Book* UnwrapBook(v8::Handle<v8::Value> bookHandle);
libxl::Book* UnwrapBook(Book* book);
libxl::Book* UnwrapBook(BookWrapper* bookWrapper);

template<typename T> _NAN_METHOD_RETURN_TYPE ThrowLibxlError(T wrappedBook) {
    NanScope();

    libxl::Book* book = UnwrapBook(wrappedBook);

    return NanThrowError(book ? book->errorMessage() : "");
}


template<typename T, typename U> bool IsSameBook(T book1, U book2) {
    libxl::Book *libxlBook1 = UnwrapBook(book1), *libxlBook2 = UnwrapBook(book2);

    return !libxlBook1 || !libxlBook2 || libxlBook1 == libxlBook2;
}


}
}


#endif // BINDINGS_UTIL
