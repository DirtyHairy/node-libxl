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
#include "font.h"
#include "api_key.h"

using namespace v8;

namespace node_libxl {


// Lifecycle


Book::Book(libxl::Book* libxlBook) :
    Wrapper<libxl::Book>(libxlBook)
{}


Book::~Book() {
    wrapped->release();
}


NAN_METHOD(Book::New) {
    NanScope();

    if (!args.IsConstructCall()) {
        NanReturnValue(NanNew(
            util::ProxyConstructor(NanNew(constructor), args)));
    }

    ArgumentHelper arguments(args);

    int32_t type = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    libxl::Book* libxlBook;

    switch (type) {
        case BOOK_TYPE_XLS:
            libxlBook = xlCreateBook();
            break;
        case BOOK_TYPE_XLSX:
            libxlBook = xlCreateXMLBook();
            break;
        default:
            CSNanThrow(Exception::TypeError(
                NanNew<String>("invalid book type")
            ));
    }

    if (!libxlBook) {
        CSNanThrow(Exception::Error(
            NanNew<String>("unknown error")));
    }

    libxlBook->setLocale("UTF-8");
    #ifdef INCLUDE_API_KEY
        libxlBook->setKey(API_KEY_NAME, API_KEY_KEY);
    #endif


    Book* book = new Book(libxlBook);
    book->Wrap(args.This());

    NanReturnValue(args.This());
}


// Wrappers


NAN_METHOD(Book::WriteSync) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value filename(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Book* libxlBook = that->GetWrapped();
    if (!libxlBook->save(*filename)) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Book::AddSheet) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Sheet* parentSheet = Sheet::Unwrap(args[1]);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (parentSheet) {
        ASSERT_SAME_BOOK(parentSheet, that);
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Sheet* libxlSheet = libxlBook->addSheet(*name,
        parentSheet ? parentSheet->GetWrapped() : NULL);

    if (!libxlSheet) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(Sheet::NewInstance(
        libxlSheet, args.This()));
}


NAN_METHOD(Book::AddCustomNumFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value description(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);
    
    libxl::Book* libxlBook = that->GetWrapped();
    int32_t format = libxlBook->addCustomNumFormat(*description);

    if (!format) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(NanNew<Number>(format));
}


NAN_METHOD(Book::AddFormat) {
    NanScope();

    Format* parentFormat = Format::Unwrap(args[0]);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (parentFormat) {
        ASSERT_SAME_BOOK(parentFormat, that);
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Format* libxlFormat = libxlBook->addFormat(
        parentFormat ? parentFormat->GetWrapped() : NULL);

    if (!libxlFormat) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(NanNew(
        Format::NewInstance(libxlFormat, args.This())));
}


NAN_METHOD(Book::AddFont) {
    NanScope();

    Font* parentFont = Font::Unwrap(args[0]);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (parentFont) {
        ASSERT_SAME_BOOK(parentFont, that);
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Font* libxlFont = libxlBook->addFont(
        parentFont ? parentFont->GetWrapped() : NULL);

    if (!libxlFont) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(NanNew(
        Font::NewInstance(libxlFont, args.This())));
}


// Init


void Book::Initialize(Handle<Object> exports) {
    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(New);
    t->SetClassName(NanNew<String>("Book"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    NODE_SET_PROTOTYPE_METHOD(t, "writeSync", WriteSync);
    NODE_SET_PROTOTYPE_METHOD(t, "addSheet", AddSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "addFormat", AddFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "addCustomNumFormat", AddCustomNumFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "addFont", AddFont);

    #ifdef INCLUDE_API_KEY
        exports->Set(NanNew<String>("apiKeyCompiledIn"), NanTrue(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #else
        exports->Set(NanNew<String>("apiKeyCompiledIn"), NanFalse(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #endif

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());
    exports->Set(NanSymbol("Book"), NanNew(constructor));

    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLS);
    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLSX);
}


}
