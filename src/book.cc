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

    int type = arguments.GetInt(0);
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
            return NanThrowTypeError("invalid book type");
    }

    if (!libxlBook) {
        return NanThrowError("unknown error");
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


NAN_METHOD(Book::LoadSync){
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value filename(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->load(*filename)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


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
    Sheet* parentSheet = arguments.GetWrapped<Sheet>(1, NULL);
    ASSERT_ARGUMENTS(arguments);


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


NAN_METHOD(Book::InsertSheet) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    String::Utf8Value name(arguments.GetString(1));
    Sheet* parentSheet = arguments.GetWrapped<Sheet>(2, NULL);
    ASSERT_ARGUMENTS(arguments);


    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);
    if (parentSheet) {
        ASSERT_SAME_BOOK(parentSheet, that);
    }

    libxl::Sheet* libxlSheet = that->GetWrapped()->insertSheet(index, *name,
        parentSheet ? parentSheet->GetWrapped() : NULL);

    if (!libxlSheet) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(Sheet::NewInstance(libxlSheet, args.This()));
}


NAN_METHOD(Book::GetSheet) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Sheet* sheet = that->GetWrapped()->getSheet(index);
    if (!sheet) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(Sheet::NewInstance(sheet, args.This()));
}


NAN_METHOD(Book::SheetType) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->sheetType(index)));
}


NAN_METHOD(Book::DelSheet) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delSheet(index)) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(args.This());
}


NAN_METHOD(Book::SheetCount) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->sheetCount()));
}


NAN_METHOD(Book::AddFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    node_libxl::Format* parentFormat =
        arguments.GetWrapped<node_libxl::Format>(0, NULL);
    ASSERT_ARGUMENTS(arguments);

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

    ArgumentHelper arguments(args);

    node_libxl::Font* parentFont =
        arguments.GetWrapped<node_libxl::Font>(0, NULL);
    ASSERT_ARGUMENTS(arguments);

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


NAN_METHOD(Book::AddCustomNumFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value description(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);
    
    libxl::Book* libxlBook = that->GetWrapped();
    int format = libxlBook->addCustomNumFormat(*description);

    if (!format) {
        return util::ThrowLibxlError(libxlBook);
    }

    NanReturnValue(NanNew<Number>(format));
}


NAN_METHOD(Book::CustomNumFormat) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    const char* formatString = that->GetWrapped()->customNumFormat(index);
    if (!formatString) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(NanNew<String>(formatString));
}


NAN_METHOD(Book::Format) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Format* format = that->GetWrapped()->format(index);
    if (!format) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(Format::NewInstance(format, args.This()));
}


NAN_METHOD(Book::FormatSize) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->formatSize()));
}


NAN_METHOD(Book::Font) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    libxl::Font* font = that->GetWrapped()->font(index);
    if (!font) {
        return util::ThrowLibxlError(that);
    }

    NanReturnValue(Font::NewInstance(font, args.This()));
}

NAN_METHOD(Book::FontSize) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->fontSize()));
}


NAN_METHOD(Book::DatePack) {
    NanScope();

    ArgumentHelper arguments(args);

    int year        = arguments.GetInt(0),
        month       = arguments.GetInt(1),
        day         = arguments.GetInt(2),
        hour        = arguments.GetInt(3, 0),
        minute      = arguments.GetInt(4, 0),
        second      = arguments.GetInt(5, 0),
        msecond     = arguments.GetInt(6, 0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Number>(that->GetWrapped()->datePack(
        year, month, day, hour, minute, second, msecond)));
}


NAN_METHOD(Book::DateUnpack) {
    NanScope();

    ArgumentHelper arguments(args);

    double value = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int year, month, day, hour, minute, second, msecond;
    if (!that->GetWrapped()->dateUnpack(
        value, &year, &month, &day, &hour, &minute, &second, &msecond))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("year"),      NanNew<Integer>(year));
    result->Set(NanNew<String>("month"),     NanNew<Integer>(month));
    result->Set(NanNew<String>("day"),       NanNew<Integer>(day));
    result->Set(NanNew<String>("hour"),      NanNew<Integer>(hour));
    result->Set(NanNew<String>("minute"),    NanNew<Integer>(minute));
    result->Set(NanNew<String>("second"),    NanNew<Integer>(second));
    result->Set(NanNew<String>("msecond"),   NanNew<Integer>(msecond));

    NanReturnValue(result);
}


NAN_METHOD(Book::ColorPack) {
    NanScope();

    ArgumentHelper arguments(args);

    int red     = arguments.GetInt(0),
        green   = arguments.GetInt(1),
        blue    = arguments.GetInt(2);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->colorPack(red, green, blue)));
}


NAN_METHOD(Book::ColorUnpack) {
    NanScope();

    ArgumentHelper arguments(args);

    int value = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    Local<Object> result = NanNew<Object>();
    int red, green, blue;

    that->GetWrapped()->colorUnpack(
        static_cast<libxl::Color>(value), &red, &green, &blue);

    result->Set(NanNew<String>("red"),   NanNew<Integer>(red));
    result->Set(NanNew<String>("green"), NanNew<Integer>(green));
    result->Set(NanNew<String>("blue"),  NanNew<Integer>(blue));

    NanReturnValue(result);
}


NAN_METHOD(Book::ActiveSheet) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->activeSheet()));
}


NAN_METHOD(Book::SetActiveSheet) {
    NanScope();

    ArgumentHelper arguments(args);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setActiveSheet(index);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::DefaultFont) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    int size;
    const char* name = that->GetWrapped()->defaultFont(&size);

    if (!name) {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = NanNew<Object>();
    result->Set(NanNew<String>("name"), NanNew<String>(name));
    result->Set(NanNew<String>("size"), NanNew<Integer>(size));

    NanReturnValue(result);
}


NAN_METHOD(Book::SetDefaultFont) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value name(arguments.GetString(0));
    int size = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDefaultFont(*name, size);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::RefR1C1) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->refR1C1()));
}

NAN_METHOD(Book::SetRefR1C1) {
    NanScope();

    ArgumentHelper arguments(args);

    bool refR1C1 = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRefR1C1(refR1C1);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::RgbMode) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->rgbMode()));
}


NAN_METHOD(Book::SetRgbMode) {
    NanScope();

    ArgumentHelper arguments(args);

    bool rgbMode = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRgbMode(rgbMode);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::BiffVersion) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Integer>(that->GetWrapped()->biffVersion()));
}


NAN_METHOD(Book::IsDate1904) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->isDate1904()));
}


NAN_METHOD(Book::SetDate1904) {
    NanScope();

    ArgumentHelper arguments(args);

    bool date1904 = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDate1904(date1904);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::IsTemplate) {
    NanScope();

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    NanReturnValue(NanNew<Boolean>(that->GetWrapped()->isTemplate()));
}


NAN_METHOD(Book::SetTemplate) {
    NanScope();

    ArgumentHelper arguments(args);

    bool isTemplate = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setTemplate(isTemplate);

    NanReturnValue(args.This());
}


NAN_METHOD(Book::SetKey) {
    NanScope();

    ArgumentHelper arguments(args);

    String::Utf8Value   name(arguments.GetString(0)),
                        key(arguments.GetString(1));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(args.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setKey(*name, *key);

    NanReturnValue(args.This());
}


// Init


void Book::Initialize(Handle<Object> exports) {
    using namespace libxl;

    NanScope();

    Local<FunctionTemplate> t = NanNew<FunctionTemplate>(New);
    t->SetClassName(NanNew<String>("Book"));
    t->InstanceTemplate()->SetInternalFieldCount(1);

    NODE_SET_PROTOTYPE_METHOD(t, "loadSync", LoadSync);
    NODE_SET_PROTOTYPE_METHOD(t, "writeSync", WriteSync);
    NODE_SET_PROTOTYPE_METHOD(t, "addSheet", AddSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "insertSheet", InsertSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "getSheet", GetSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "sheetType", SheetType);
    NODE_SET_PROTOTYPE_METHOD(t, "delSheet", DelSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "sheetCount", SheetCount);
    NODE_SET_PROTOTYPE_METHOD(t, "addFormat", AddFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "addFont", AddFont);
    NODE_SET_PROTOTYPE_METHOD(t, "addCustomNumFormat", AddCustomNumFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "customNumFormat", CustomNumFormat);
    NODE_SET_PROTOTYPE_METHOD(t, "format", Format);
    NODE_SET_PROTOTYPE_METHOD(t, "formatSize", FormatSize);
    NODE_SET_PROTOTYPE_METHOD(t, "font", Font);
    NODE_SET_PROTOTYPE_METHOD(t, "fontSize", FontSize);
    NODE_SET_PROTOTYPE_METHOD(t, "datePack", DatePack);
    NODE_SET_PROTOTYPE_METHOD(t, "dateUnpack", DateUnpack);
    NODE_SET_PROTOTYPE_METHOD(t, "colorPack", ColorPack);
    NODE_SET_PROTOTYPE_METHOD(t, "colorUnpack", ColorUnpack);
    NODE_SET_PROTOTYPE_METHOD(t, "activeSheet", ActiveSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "setActiveSheet", SetActiveSheet);
    NODE_SET_PROTOTYPE_METHOD(t, "defaultFont", DefaultFont);
    NODE_SET_PROTOTYPE_METHOD(t, "setDefaultFont", SetDefaultFont);
    NODE_SET_PROTOTYPE_METHOD(t, "refR1C1", RefR1C1);
    NODE_SET_PROTOTYPE_METHOD(t, "setRefR1C1", SetRefR1C1);
    NODE_SET_PROTOTYPE_METHOD(t, "rgbMode", RgbMode);
    NODE_SET_PROTOTYPE_METHOD(t, "setRgbMode", SetRgbMode);
    NODE_SET_PROTOTYPE_METHOD(t, "biffVersion", BiffVersion);
    NODE_SET_PROTOTYPE_METHOD(t, "isDate1904", IsDate1904);
    NODE_SET_PROTOTYPE_METHOD(t, "setDate1904", SetDate1904);
    NODE_SET_PROTOTYPE_METHOD(t, "isTemplate", IsTemplate);
    NODE_SET_PROTOTYPE_METHOD(t, "setTemplate", SetTemplate);
    NODE_SET_PROTOTYPE_METHOD(t, "setKey", SetKey);

    #ifdef INCLUDE_API_KEY
        exports->Set(NanNew<String>("apiKeyCompiledIn"), NanTrue(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #else
        exports->Set(NanNew<String>("apiKeyCompiledIn"), NanFalse(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #endif

    t->ReadOnlyPrototype();
    NanAssignPersistent(constructor, t->GetFunction());
    exports->Set(NanNew<String>("Book"), NanNew(constructor));

    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLS);
    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLSX);
    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_SHEET);
    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_CHART);
    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_UNKNOWN);
}


}
