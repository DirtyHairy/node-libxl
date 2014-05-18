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

#ifndef BINDINGS_SHEET_H
#define BINDINGS_SHEET_H

#include "common.h"
#include "wrapper.h"
#include "book_wrapper.h"

namespace node_libxl {


class Sheet : public Wrapper<libxl::Sheet> , public BookWrapper

{
    public:

        Sheet(libxl::Sheet* sheet, v8::Handle<v8::Value> book);

        static void Initialize(v8::Handle<v8::Object> exports);
        
        static Sheet* Unwrap(v8::Handle<v8::Value> object) {
            return Wrapper<libxl::Sheet>::Unwrap<Sheet>(object);
        }

        static v8::Handle<v8::Object> NewInstance(
            libxl::Sheet* sheet,
            v8::Handle<v8::Value> book
        );

    protected:

        static NAN_METHOD(CellType);
        static NAN_METHOD(IsFormula);
        static NAN_METHOD(CellFormat);
        static NAN_METHOD(SetCellFormat);
        static NAN_METHOD(ReadStr);
        static NAN_METHOD(WriteStr);
        static NAN_METHOD(ReadNum);
        static NAN_METHOD(WriteNum);
        static NAN_METHOD(ReadBool);
        static NAN_METHOD(WriteBool);
        static NAN_METHOD(ReadBlank);
        static NAN_METHOD(WriteBlank);
        static NAN_METHOD(ReadFormula);
        static NAN_METHOD(WriteFormula);
        static NAN_METHOD(ReadComment);
        static NAN_METHOD(WriteComment);
        // TODO IsDate
        static NAN_METHOD(ReadError);
        static NAN_METHOD(ColWidth);
        static NAN_METHOD(RowHeight);
        static NAN_METHOD(SetCol);
        static NAN_METHOD(SetRow);
        static NAN_METHOD(RowHidden);
        static NAN_METHOD(SetRowHidden);
        static NAN_METHOD(ColHidden);
        static NAN_METHOD(SetColHidden);
        static NAN_METHOD(SetMerge);

    private:

        Sheet(const Sheet&);
        const Sheet& operator=(const Sheet&);
};


}

#endif // BINDINGS_SHEET_H
