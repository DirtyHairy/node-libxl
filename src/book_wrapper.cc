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

#include "book_wrapper.h"

#include <cstdlib>
#include <iostream>

using namespace v8;

namespace node_libxl {

    BookHolder::BookHolder(Local<Value> bookHandle) {
        if (!Book::FromJS(bookHandle)) {
            std::cerr << "libxl bindings: internal error: handle is not a book instance"
                      << std::endl;

            abort();
        }

        this->bookHandle.Reset(bookHandle);
    }

    BookHolder::~BookHolder() { bookHandle.Reset(); }

    v8::Local<v8::Value> BookHolder::GetBookHandle() {
        Nan::EscapableHandleScope scope;

        return scope.Escape(Nan::New(bookHandle));
    }

    Book* BookHolder::GetBook() {
        Nan::HandleScope scope;

        return Book::FromJS(Nan::New(bookHandle));
    }

}  // namespace node_libxl
