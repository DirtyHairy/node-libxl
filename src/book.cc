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

#include "book.h"
#include "argument_helper.h"
#include "assert.h"
#include "util.h"
#include "sheet.h"
#include "format.h"

using namespace v8;

namespace node_libxl {


// Lifecycle


Book::Book(libxl::Book* libxlBook) :
    Wrapper<libxl::Book>(libxlBook)
{}


Book::~Book() {
    wrapped->release();
}


Handle<Value> Book::New(const Arguments& arguments) {
    HandleScope scope;

    if (!arguments.IsConstructCall()) {
        return scope.Close(util::ProxyConstructor(constructor, arguments));
    }

    ArgumentHelper args(arguments);

    int32_t type = args.GetInt(0);
    ASSERT_ARGUMENTS(args);

    libxl::Book* libxlBook;

    switch (type) {
        case BOOK_TYPE_XLS:
            libxlBook = xlCreateBook();
            break;
        case BOOK_TYPE_XLSX:
            libxlBook = xlCreateXMLBook();
            break;
        default:
            return ThrowException(Exception::TypeError(String::New(
                "invalid book type")));
    }

    if (!libxlBook) {
        return ThrowException(Exception::Error(String::New("unknown error")));
    }

    libxlBook->setLocale("UTF-8");
    Book* book = new Book(libxlBook);
    book->Wrap(arguments.This());

    return arguments.This();
}


// Wrappers


Handle<Value> Book::WriteSync(const Arguments& arguments) {
    HandleScope scope;

    ArgumentHelper args(arguments);

    String::Utf8Value filename(args.GetString(0));
    ASSERT_ARGUMENTS(args);

    Book* that = Unwrap(arguments.This());
    ASSERT_THIS(that);

    libxl::Book* libxlBook = that->GetWrapped();
    if (!libxlBook->save(*filename)) {
        return util::ThrowLibxlError(libxlBook);
    }

    return scope.Close(arguments.This());
}


Handle<Value> Book::AddSheet(const Arguments& arguments) {
    HandleScope scope;

    ArgumentHelper args(arguments);

    String::Utf8Value name(args.GetString(0));
    ASSERT_ARGUMENTS(args);

    Sheet* parentSheet = Sheet::Unwrap(arguments[1]);

    Book* that = Unwrap(arguments.This());
    ASSERT_THIS(that);
    if (parentSheet) {
        ASSERT_SAME_BOOK(parentSheet->GetBookHandle(), that);
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Sheet* libxlSheet = libxlBook->addSheet(*name,
        parentSheet ? parentSheet->GetWrapped() : NULL);

    if (!libxlSheet) {
        return util::ThrowLibxlError(libxlBook);
    }

    return scope.Close(Sheet::NewInstance(libxlSheet, arguments.This()));
}


Handle<Value> Book::AddFormat(const Arguments& arguments) {
    HandleScope scope;

    Format* parentFormat = Format::Unwrap(arguments[0]);

    Book* that = Unwrap(arguments.This());
    ASSERT_THIS(that);

    if (parentFormat) {
        ASSERT_SAME_BOOK(parentFormat->GetBookHandle(), that);
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Format* libxlFormat = libxlBook->addFormat(
        parentFormat ? parentFormat->GetWrapped() : NULL);

    if (!libxlFormat) {
        return util::ThrowLibxlError(libxlBook);
    }

    return scope.Close(Format::NewInstance(libxlFormat, arguments.This()));
}


// Init


void Book::Initialize(Handle<Object> exports) {
    HandleScope scope;

    Local<FunctionTemplate> t = FunctionTemplate::New(New);
    t->SetClassName(String::NewSymbol("Book"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    NODE_SET_PROTOTYPE_METHOD(t, "writeSync", WriteSync);
    NODE_SET_PROTOTYPE_METHOD(t, "addSheet", AddSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "addFormat", AddFormat);

    t->ReadOnlyPrototype();
    constructor = Persistent<Function>::New(t->GetFunction());
    exports->Set(String::NewSymbol("Book"), constructor);

    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLS);
    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLSX);
}


}
