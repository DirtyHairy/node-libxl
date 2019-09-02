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

#ifndef BINDINGS_BOOK
#define BINDINGS_BOOK

#include "common.h"
#include "wrapper.h"

namespace node_libxl {


enum {
    BOOK_TYPE_XLS,
    BOOK_TYPE_XLSX
};


class Book : public Wrapper<libxl::Book> {
    public:

        Book(libxl::Book* libxlBook);
        ~Book();

        void StartAsync();
        void StopAsync();
        bool AsyncPending();

        static void Initialize(v8::Local<v8::Object> exports);

        static Book* Unwrap(v8::Local<v8::Value> object) {
            return Wrapper<libxl::Book>::Unwrap<Book>(object);
        }

    protected:

        static NAN_METHOD(New);

        static NAN_METHOD(LoadSync);
        static NAN_METHOD(Load);
        static NAN_METHOD(WriteSync);
        static NAN_METHOD(Write);
        static NAN_METHOD(WriteRawSync);
        static NAN_METHOD(WriteRaw);
        static NAN_METHOD(LoadRawSync);
        static NAN_METHOD(LoadRaw);
        static NAN_METHOD(AddSheet);
        static NAN_METHOD(InsertSheet);
        static NAN_METHOD(GetSheet);
        static NAN_METHOD(SheetType);
        static NAN_METHOD(DelSheet);
        static NAN_METHOD(SheetCount);
        static NAN_METHOD(AddFormat);
        static NAN_METHOD(AddFont);
        static NAN_METHOD(AddCustomNumFormat);
        static NAN_METHOD(CustomNumFormat);
        static NAN_METHOD(Format);
        static NAN_METHOD(FormatSize);
        static NAN_METHOD(Font);
        static NAN_METHOD(FontSize);
        static NAN_METHOD(DatePack);
        static NAN_METHOD(DateUnpack);
        static NAN_METHOD(ColorPack);
        static NAN_METHOD(ColorUnpack);
        static NAN_METHOD(ActiveSheet);
        static NAN_METHOD(SetActiveSheet);
        static NAN_METHOD(PictureSize);
        static NAN_METHOD(GetPicture);
        static NAN_METHOD(GetPictureAsync);
        static NAN_METHOD(AddPicture);
        static NAN_METHOD(AddPictureAsync);
        static NAN_METHOD(DefaultFont);
        static NAN_METHOD(SetDefaultFont);
        static NAN_METHOD(RefR1C1);
        static NAN_METHOD(SetRefR1C1);
        static NAN_METHOD(RgbMode);
        static NAN_METHOD(SetRgbMode);
        static NAN_METHOD(BiffVersion);
        static NAN_METHOD(IsDate1904);
        static NAN_METHOD(SetDate1904);
        static NAN_METHOD(IsTemplate);
        static NAN_METHOD(SetTemplate);
        static NAN_METHOD(SetKey);

    private:

        Book(const Book&);
        const Book& operator=(const Book&);

        bool asyncPending;
};


}

#include "libxl.h"

#endif // BINDINGS_BOOK
