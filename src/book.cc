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

#include <cstring>

#include "book.h"
#include "argument_helper.h"
#include "assert.h"
#include "util.h"
#include "sheet.h"
#include "format.h"
#include "font.h"
#include "api_key.h"
#include "async_worker.h"
#include "string_copy.h"
#include "buffer_copy.h"

using namespace v8;

namespace node_libxl {


// Lifecycle


Book::Book(libxl::Book* libxlBook) :
    Wrapper<libxl::Book>(libxlBook),
    asyncPending(false)
{}


Book::~Book() {
    wrapped->release();
}


NAN_METHOD(Book::New) {
    Nan::HandleScope scope;

    if (!info.IsConstructCall()) {
        info.GetReturnValue().Set(
            util::ProxyConstructor(Nan::New(constructor), info));
    }

    ArgumentHelper arguments(info);

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
            return Nan::ThrowTypeError("invalid book type");
    }

    if (!libxlBook) {
        return Nan::ThrowError("unknown error");
    }

    libxlBook->setLocale("UTF-8");
    #ifdef INCLUDE_API_KEY
        libxlBook->setKey(API_KEY_NAME, API_KEY_KEY);
    #endif

    Book* book = new Book(libxlBook);
    book->Wrap(info.This());

    info.GetReturnValue().Set(info.This());
}


// Async guard


void Book::StartAsync() {
    asyncPending = true;
}


void Book::StopAsync() {
    asyncPending = false;
}


bool Book::AsyncPending() {
    return asyncPending;
}


// Implementation


NAN_METHOD(Book::LoadSync){
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value filename(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->load(*filename)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::Load) {
    class Worker : public AsyncWorker<Book> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that, Handle<Value> filename) :
                AsyncWorker<Book>(callback, that),
                filename(filename)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->load(*filename)) {
                    RaiseLibxlError();
                }
            }

        private:
            StringCopy filename;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Local<Value> filename = arguments.GetString(0);
    Local<Function> callback = arguments.GetFunction(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), filename));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::WriteSync) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value filename(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Book* libxlBook = that->GetWrapped();
    if (!libxlBook->save(*filename)) {
        return util::ThrowLibxlError(libxlBook);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::Write) {
    class Worker : public AsyncWorker<Book> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that, Handle<Value> filename) :
                AsyncWorker<Book>(callback, that),
                filename(filename)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->save(*filename)) {
                    RaiseLibxlError();
                }
            }

        private:
            StringCopy filename;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Local<Value> filename = arguments.GetString(0);
    Local<Function> callback = arguments.GetFunction(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), filename));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::WriteRawSync) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    const char* data;
    unsigned size;

    if (!that->GetWrapped()->saveRaw(&data, &size)) {
        return util::ThrowLibxlError(that);
    }

    char* buffer = new char[size];
    memcpy(buffer, data, size);

    info.GetReturnValue().Set(Nan::NewBuffer(buffer, size).ToLocalChecked());
}


NAN_METHOD(Book::WriteRaw) {
    class Worker : public AsyncWorker<Book> {
        public:
            Worker(Nan::Callback *callback, Local<Object> that) :
                AsyncWorker<Book>(callback, that)
            {}

            virtual void Execute() {
                const char* data;

                if (!that->GetWrapped()->saveRaw(&data, &size)) {
                    RaiseLibxlError();
                } else {
                    buffer = new char[size];
                    memcpy(buffer, data, size);
                }
            }

            virtual void HandleOKCallback() {
                Nan::HandleScope scope;

                Local<Value> argv[] = {
                    Nan::Undefined(),
                    Nan::NewBuffer(buffer, size).ToLocalChecked()
                };
                callback->Call(2, argv);
            }

        private:
            char* buffer;
            unsigned size;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Local<Function> callback = arguments.GetFunction(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This()));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::LoadRawSync) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Handle<Value> buffer = arguments.GetBuffer(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->loadRaw(
        node::Buffer::Data(buffer), node::Buffer::Length(buffer)))
    {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::LoadRaw) {
    class Worker : public AsyncWorker<Book> {
        public:
            Worker(Nan::Callback *callback, Local<Object> that,
                    Handle<Value> buffer) :
                AsyncWorker<Book>(callback, that),
                buffer(buffer)
            {}

            virtual void Execute() {
                if (!that->GetWrapped()->loadRaw(*buffer, buffer.GetSize())) {
                    RaiseLibxlError();
                }
            }

        private:
            BufferCopy buffer;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Local<Value> buffer = arguments.GetBuffer(0);
    Local<Function> callback = arguments.GetFunction(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(
        new Nan::Callback(callback), info.This(), buffer));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::AddSheet) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    Sheet* parentSheet = arguments.GetWrapped<Sheet>(1, NULL);
    ASSERT_ARGUMENTS(arguments);


    Book* that = Unwrap(info.This());
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

    info.GetReturnValue().Set(Sheet::NewInstance(
        libxlSheet, info.This()));
}


NAN_METHOD(Book::InsertSheet) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    String::Utf8Value name(arguments.GetString(1));
    Sheet* parentSheet = arguments.GetWrapped<Sheet>(2, NULL);
    ASSERT_ARGUMENTS(arguments);


    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);
    if (parentSheet) {
        ASSERT_SAME_BOOK(parentSheet, that);
    }

    libxl::Sheet* libxlSheet = that->GetWrapped()->insertSheet(index, *name,
        parentSheet ? parentSheet->GetWrapped() : NULL);

    if (!libxlSheet) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Sheet::NewInstance(libxlSheet, info.This()));
}


NAN_METHOD(Book::GetSheet) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Sheet* sheet = that->GetWrapped()->getSheet(index);
    if (!sheet) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Sheet::NewInstance(sheet, info.This()));
}


NAN_METHOD(Book::SheetType) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->sheetType(index)));
}


NAN_METHOD(Book::DelSheet) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (!that->GetWrapped()->delSheet(index)) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::SheetCount) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->sheetCount()));
}


NAN_METHOD(Book::AddFormat) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    node_libxl::Format* parentFormat =
        arguments.GetWrapped<node_libxl::Format>(0, NULL);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (parentFormat && parentFormat->GetBook()->AsyncPending()) {
        return Nan::ThrowError("async operation pending on parent");
    }

    libxl::Book* libxlBook = that->GetWrapped();
    libxl::Format* libxlFormat = libxlBook->addFormat(
        parentFormat ? parentFormat->GetWrapped() : NULL);

    if (!libxlFormat) {
        return util::ThrowLibxlError(libxlBook);
    }

    info.GetReturnValue().Set(
        Format::NewInstance(libxlFormat, info.This()));
}


NAN_METHOD(Book::AddFont) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    node_libxl::Font* parentFont =
        arguments.GetWrapped<node_libxl::Font>(0, NULL);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
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

    info.GetReturnValue().Set(
        Font::NewInstance(libxlFont, info.This()));
}


NAN_METHOD(Book::AddCustomNumFormat) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value description(arguments.GetString(0));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Book* libxlBook = that->GetWrapped();
    int format = libxlBook->addCustomNumFormat(*description);

    if (!format) {
        return util::ThrowLibxlError(libxlBook);
    }

    info.GetReturnValue().Set(Nan::New<Number>(format));
}


NAN_METHOD(Book::CustomNumFormat) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    const char* formatString = that->GetWrapped()->customNumFormat(index);
    if (!formatString) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Nan::New<String>(formatString).ToLocalChecked());
}


NAN_METHOD(Book::Format) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Format* format = that->GetWrapped()->format(index);
    if (!format) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Format::NewInstance(format, info.This()));
}


NAN_METHOD(Book::FormatSize) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->formatSize()));
}


NAN_METHOD(Book::Font) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    libxl::Font* font = that->GetWrapped()->font(index);
    if (!font) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Font::NewInstance(font, info.This()));
}

NAN_METHOD(Book::FontSize) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->fontSize()));
}


NAN_METHOD(Book::DatePack) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int year        = arguments.GetInt(0),
        month       = arguments.GetInt(1),
        day         = arguments.GetInt(2),
        hour        = arguments.GetInt(3, 0),
        minute      = arguments.GetInt(4, 0),
        second      = arguments.GetInt(5, 0),
        msecond     = arguments.GetInt(6, 0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->datePack(
        year, month, day, hour, minute, second, msecond)));
}


NAN_METHOD(Book::DateUnpack) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    double value = arguments.GetDouble(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int year, month, day, hour, minute, second, msecond;
    if (!that->GetWrapped()->dateUnpack(
        value, &year, &month, &day, &hour, &minute, &second, &msecond))
    {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("year").ToLocalChecked(),      Nan::New<Integer>(year));
    result->Set(Nan::New<String>("month").ToLocalChecked(),     Nan::New<Integer>(month));
    result->Set(Nan::New<String>("day").ToLocalChecked(),       Nan::New<Integer>(day));
    result->Set(Nan::New<String>("hour").ToLocalChecked(),      Nan::New<Integer>(hour));
    result->Set(Nan::New<String>("minute").ToLocalChecked(),    Nan::New<Integer>(minute));
    result->Set(Nan::New<String>("second").ToLocalChecked(),    Nan::New<Integer>(second));
    result->Set(Nan::New<String>("msecond").ToLocalChecked(),   Nan::New<Integer>(msecond));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Book::ColorPack) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int red     = arguments.GetInt(0),
        green   = arguments.GetInt(1),
        blue    = arguments.GetInt(2);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->colorPack(red, green, blue)));
}


NAN_METHOD(Book::ColorUnpack) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int value = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Local<Object> result = Nan::New<Object>();
    int red, green, blue;

    that->GetWrapped()->colorUnpack(
        static_cast<libxl::Color>(value), &red, &green, &blue);

    result->Set(Nan::New<String>("red").ToLocalChecked(),   Nan::New<Integer>(red));
    result->Set(Nan::New<String>("green").ToLocalChecked(), Nan::New<Integer>(green));
    result->Set(Nan::New<String>("blue").ToLocalChecked(),  Nan::New<Integer>(blue));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Book::ActiveSheet) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->activeSheet()));
}


NAN_METHOD(Book::SetActiveSheet) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setActiveSheet(index);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::PictureSize) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->pictureSize()));
}


NAN_METHOD(Book::GetPicture) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    const char* data;
    unsigned size;
    libxl::PictureType pictureType = that->GetWrapped()
        ->getPicture(index, &data, &size);

    if (pictureType == libxl::PICTURETYPE_ERROR) {
        return util::ThrowLibxlError(that);
    }

    char* buffer = new char[size];
    memcpy(buffer, data, size);

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("type").ToLocalChecked(), Nan::New<Integer>(pictureType));
    result->Set(Nan::New<String>("data").ToLocalChecked(), Nan::NewBuffer(buffer, size).ToLocalChecked());

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Book::GetPictureAsync) {
    class Worker : public AsyncWorker<Book> {
        public:
            Worker(Nan::Callback* callback, Local<Object> that, int index) :
                AsyncWorker<Book>(callback, that),
                index(index)
            {}

            virtual void Execute() {
                const char* data;

                pictureType = that->GetWrapped()->getPicture(index, &data, &size);
                if (pictureType == libxl::PICTURETYPE_ERROR) {
                    RaiseLibxlError();
                } else {
                    buffer = new char[size];
                    memcpy(buffer, data, size);
                }
            }

            virtual void HandleOKCallback() {
                Nan::HandleScope scope;

                Local<Value> argv[] = {
                    Nan::Undefined(),
                    Nan::New<Integer>(pictureType),
                    Nan::NewBuffer(buffer, size).ToLocalChecked()
                };

                callback->Call(3, argv);
            }

        private:
            int index, pictureType;
            char* buffer;
            unsigned size;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    int index = arguments.GetInt(0);
    Local<Function> callback = arguments.GetFunction(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    Nan::AsyncQueueWorker(new Worker(new Nan::Callback(callback), info.This(), index));

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::AddPicture) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int index;

    if (info[0]->IsString()) {

        String::Utf8Value filename(arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        index = that->GetWrapped()->addPicture(*filename);

    } else if (node::Buffer::HasInstance(info[0])) {

        Handle<Value> buffer = arguments.GetBuffer(0);
        ASSERT_ARGUMENTS(arguments);

        index = that->GetWrapped()->addPicture2(
            node::Buffer::Data(buffer), node::Buffer::Length(buffer));

    } else {
        return Nan::ThrowTypeError("string or buffer required as argument 0");
    }

    if (index == -1) {
        return util::ThrowLibxlError(that);
    }

    info.GetReturnValue().Set(Nan::New<Integer>(index));
}


NAN_METHOD(Book::AddPictureAsync) {
    class FileWorker : public AsyncWorker<Book> {
        public:
            FileWorker(Nan::Callback* callback, Local<Object> that,
                    Handle<Value> filename) :
                AsyncWorker<Book>(callback, that),
                filename(filename)
            {}

            virtual void Execute() {
                index = that->GetWrapped()->addPicture(*filename);
                if (index == -1) {
                    RaiseLibxlError();
                }
            }

            virtual void HandleOKCallback() {
                Nan::HandleScope scope;

                Local<Value> argv[] = {
                    Nan::Undefined(),
                    Nan::New<Integer>(index)
                };

                callback->Call(2, argv);
            }

        private:
            StringCopy filename;
            int index;
    };

    class BufferWorker : public AsyncWorker<Book> {
        public:
            BufferWorker(Nan::Callback* callback, Local<Object> that,
                    Handle<Value> buffer) :
                AsyncWorker<Book>(callback, that),
                buffer(buffer)
            {}

            virtual void Execute() {
                index = that->GetWrapped()->addPicture2(*buffer, buffer.GetSize());
                if (index == -1) {
                    RaiseLibxlError();
                }
            }

            virtual void HandleOKCallback() {
                Nan::HandleScope scope;

                Local<Value> argv[] = {
                    Nan::Undefined(),
                    Nan::New<Integer>(index)
                };

                callback->Call(2, argv);
            }

        private:
            BufferCopy buffer;
            int index;
    };

    Nan::HandleScope scope;

    ArgumentHelper arguments(info);
    Local<Function> callback = arguments.GetFunction(1);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    if (info[0]->IsString()) {

        Handle<Value> filename = arguments.GetString(0);
        ASSERT_ARGUMENTS(arguments);

        Nan::AsyncQueueWorker(new FileWorker(
            new Nan::Callback(callback), info.This(), filename));

    } else if (node::Buffer::HasInstance(info[0])) {

        Handle<Value> buffer = arguments.GetBuffer(0);
        ASSERT_ARGUMENTS(arguments);

        Nan::AsyncQueueWorker(new BufferWorker(
            new Nan::Callback(callback), info.This(), buffer));

    } else {
        return Nan::ThrowTypeError("string or buffer required as argument 0");
    }

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::DefaultFont) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    int size;
    const char* name = that->GetWrapped()->defaultFont(&size);

    if (!name) {
        return util::ThrowLibxlError(that);
    }

    Local<Object> result = Nan::New<Object>();
    result->Set(Nan::New<String>("name").ToLocalChecked(), Nan::New<String>(name).ToLocalChecked());
    result->Set(Nan::New<String>("size").ToLocalChecked(), Nan::New<Integer>(size));

    info.GetReturnValue().Set(result);
}


NAN_METHOD(Book::SetDefaultFont) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value name(arguments.GetString(0));
    int size = arguments.GetInt(1);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDefaultFont(*name, size);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::RefR1C1) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->refR1C1()));
}

NAN_METHOD(Book::SetRefR1C1) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool refR1C1 = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRefR1C1(refR1C1);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::RgbMode) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->rgbMode()));
}


NAN_METHOD(Book::SetRgbMode) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool rgbMode = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setRgbMode(rgbMode);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::BiffVersion) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Integer>(that->GetWrapped()->biffVersion()));
}


NAN_METHOD(Book::IsDate1904) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isDate1904()));
}


NAN_METHOD(Book::SetDate1904) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool date1904 = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setDate1904(date1904);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::IsTemplate) {
    Nan::HandleScope scope;

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->isTemplate()));
}


NAN_METHOD(Book::SetTemplate) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    bool isTemplate = arguments.GetBoolean(0, true);
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setTemplate(isTemplate);

    info.GetReturnValue().Set(info.This());
}


NAN_METHOD(Book::SetKey) {
    Nan::HandleScope scope;

    ArgumentHelper arguments(info);

    String::Utf8Value   name(arguments.GetString(0)),
                        key(arguments.GetString(1));
    ASSERT_ARGUMENTS(arguments);

    Book* that = Unwrap(info.This());
    ASSERT_THIS(that);

    that->GetWrapped()->setKey(*name, *key);

    info.GetReturnValue().Set(info.This());
}


// Init


void Book::Initialize(Handle<Object> exports) {
    using namespace libxl;

    Nan::HandleScope scope;

    Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
    t->SetClassName(Nan::New<String>("Book").ToLocalChecked());
    t->InstanceTemplate()->SetInternalFieldCount(1);

    Nan::SetPrototypeMethod(t, "loadSync", LoadSync);
    Nan::SetPrototypeMethod(t, "load", Load);
    Nan::SetPrototypeMethod(t, "writeSync", WriteSync);
    Nan::SetPrototypeMethod(t, "saveSync", WriteSync);
    Nan::SetPrototypeMethod(t, "write", Write);
    Nan::SetPrototypeMethod(t, "save", Write);
    Nan::SetPrototypeMethod(t, "loadRawSync", LoadRawSync);
    Nan::SetPrototypeMethod(t, "loadRaw", LoadRaw);
    Nan::SetPrototypeMethod(t, "writeRawSync", WriteRawSync);
    Nan::SetPrototypeMethod(t, "saveRawSync", WriteRawSync);
    Nan::SetPrototypeMethod(t, "writeRaw", WriteRaw);
    Nan::SetPrototypeMethod(t, "saveRaw", WriteRaw);
    Nan::SetPrototypeMethod(t, "addSheet", AddSheet);
    Nan::SetPrototypeMethod(t, "insertSheet", InsertSheet);
    Nan::SetPrototypeMethod(t, "getSheet", GetSheet);
    Nan::SetPrototypeMethod(t, "sheetType", SheetType);
    Nan::SetPrototypeMethod(t, "delSheet", DelSheet);
    Nan::SetPrototypeMethod(t, "sheetCount", SheetCount);
    Nan::SetPrototypeMethod(t, "addFormat", AddFormat);
    Nan::SetPrototypeMethod(t, "addFont", AddFont);
    Nan::SetPrototypeMethod(t, "addCustomNumFormat", AddCustomNumFormat);
    Nan::SetPrototypeMethod(t, "customNumFormat", CustomNumFormat);
    Nan::SetPrototypeMethod(t, "format", Format);
    Nan::SetPrototypeMethod(t, "formatSize", FormatSize);
    Nan::SetPrototypeMethod(t, "font", Font);
    Nan::SetPrototypeMethod(t, "fontSize", FontSize);
    Nan::SetPrototypeMethod(t, "datePack", DatePack);
    Nan::SetPrototypeMethod(t, "dateUnpack", DateUnpack);
    Nan::SetPrototypeMethod(t, "colorPack", ColorPack);
    Nan::SetPrototypeMethod(t, "colorUnpack", ColorUnpack);
    Nan::SetPrototypeMethod(t, "activeSheet", ActiveSheet);
    Nan::SetPrototypeMethod(t, "setActiveSheet", SetActiveSheet);
    Nan::SetPrototypeMethod(t, "pictureSize", PictureSize);
    Nan::SetPrototypeMethod(t, "getPicture", GetPicture);
    Nan::SetPrototypeMethod(t, "getPictureAsync", GetPictureAsync);
    Nan::SetPrototypeMethod(t, "addPicture", AddPicture);
    Nan::SetPrototypeMethod(t, "addPictureAsync", AddPictureAsync);
    Nan::SetPrototypeMethod(t, "defaultFont", DefaultFont);
    Nan::SetPrototypeMethod(t, "setDefaultFont", SetDefaultFont);
    Nan::SetPrototypeMethod(t, "refR1C1", RefR1C1);
    Nan::SetPrototypeMethod(t, "setRefR1C1", SetRefR1C1);
    Nan::SetPrototypeMethod(t, "rgbMode", RgbMode);
    Nan::SetPrototypeMethod(t, "setRgbMode", SetRgbMode);
    Nan::SetPrototypeMethod(t, "biffVersion", BiffVersion);
    Nan::SetPrototypeMethod(t, "isDate1904", IsDate1904);
    Nan::SetPrototypeMethod(t, "setDate1904", SetDate1904);
    Nan::SetPrototypeMethod(t, "isTemplate", IsTemplate);
    Nan::SetPrototypeMethod(t, "setTemplate", SetTemplate);
    Nan::SetPrototypeMethod(t, "setKey", SetKey);

    #ifdef INCLUDE_API_KEY
        CSNanObjectSetWithAttributes(exports, Nan::New<String>("apiKeyCompiledIn").ToLocalChecked(), Nan::True(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #else
        CSNanObjectSetWithAttributes(exports, Nan::New<String>("apiKeyCompiledIn").ToLocalChecked(), Nan::False(),
            static_cast<PropertyAttribute>(ReadOnly|DontDelete));
    #endif

    t->ReadOnlyPrototype();
    constructor.Reset(t->GetFunction());
    exports->Set(Nan::New<String>("Book").ToLocalChecked(), Nan::New(constructor));

    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLS);
    NODE_DEFINE_CONSTANT(exports, BOOK_TYPE_XLSX);

    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_SHEET);
    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_CHART);
    NODE_DEFINE_CONSTANT(exports, SHEETTYPE_UNKNOWN);

    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_PNG);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_JPEG);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_WMF);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_DIB);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_EMF);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_PICT);
    NODE_DEFINE_CONSTANT(exports, PICTURETYPE_TIFF);;
}


}
