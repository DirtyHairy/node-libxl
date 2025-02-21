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

#ifndef BINDINGS_BOOK_WRAPPER_H
#define BINDINGS_BOOK_WRAPPER_H

#include "book.h"
#include "common.h"

namespace node_libxl {

    class BookHolder {
       public:
        BookHolder(v8::Local<v8::Value> bookHandle);
        ~BookHolder();

        v8::Local<v8::Value> GetBookHandle();
        Book* GetBook();

       protected:
        Nan::Persistent<v8::Value> bookHandle;

        // We need to template this in order to unwrap the correct object
        // pointer
        template <typename T>
        static void Initialize(v8::Local<v8::FunctionTemplate> constructor);

       private:
        // We need to template this in order to unwrap the correct object
        // pointer
        template <typename T>
        static NAN_GETTER(BookAccessor);

        BookHolder();
        BookHolder(const BookHolder&);
        const BookHolder& operator=(const BookHolder&);
    };

    // Implementation

    template <typename T>
    void BookHolder::Initialize(v8::Local<v8::FunctionTemplate> constructor) {
        v8::Local<v8::ObjectTemplate> instanceTemplate = constructor->InstanceTemplate();

        Nan::SetAccessor(instanceTemplate, Nan::New<v8::String>("book").ToLocalChecked(),
                         BookAccessor<T>, NULL, v8::Local<v8::Value>(), v8::DEFAULT,
                         static_cast<v8::PropertyAttribute>(v8::ReadOnly | v8::DontDelete));
    }

    template <typename T>
    NAN_GETTER(BookHolder::BookAccessor) {
        Nan::HandleScope scope;

        BookHolder* bookWrapper =
            dynamic_cast<BookHolder*>(Nan::ObjectWrap::Unwrap<T>(info.This()));

        if (bookWrapper) {
            info.GetReturnValue().Set(Nan::New(bookWrapper->bookHandle));
        } else {
            return;
        }
    }

}  // namespace node_libxl

#endif  // BINDINGS_BOOK_WRAPPER_H
