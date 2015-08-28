/*
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

#ifndef BINDINGS_FORMAT_H
#define BINDINGS_FORMAT_H

#include "common.h"
#include "wrapper.h"
#include "book_wrapper.h"

namespace node_libxl {


class Format : public Wrapper<libxl::Format>, public BookWrapper
{
    public:

        Format(libxl::Format* format, v8::Handle<v8::Value> book);

        static void Initialize(v8::Handle<v8::Object> exports);
        
        static Format* Unwrap(v8::Handle<v8::Value> object) {
            return Wrapper<libxl::Format>::Unwrap<Format>(object);
        }

        static v8::Local<v8::Object> NewInstance(
            libxl::Format* format,
            v8::Handle<v8::Value> book
        );

    protected:

        static NAN_METHOD(Font);
        static NAN_METHOD(SetFont);
        static NAN_METHOD(NumFormat);
        static NAN_METHOD(SetNumFormat);
        static NAN_METHOD(AlignH);
        static NAN_METHOD(SetAlignH);
        static NAN_METHOD(AlignV);
        static NAN_METHOD(SetAlignV);
        static NAN_METHOD(GetWrap);
        static NAN_METHOD(SetWrap);
        static NAN_METHOD(Rotation);
        static NAN_METHOD(SetRotation);
        static NAN_METHOD(Indent);
        static NAN_METHOD(SetIndent);
        static NAN_METHOD(ShrinkToFit);
        static NAN_METHOD(SetShrinkToFit);
        static NAN_METHOD(SetBorder);
        static NAN_METHOD(SetBorderColor);
        static NAN_METHOD(BorderLeft);
        static NAN_METHOD(SetBorderLeft);
        static NAN_METHOD(BorderRight);
        static NAN_METHOD(SetBorderRight);
        static NAN_METHOD(BorderTop);
        static NAN_METHOD(SetBorderTop);
        static NAN_METHOD(BorderBottom);
        static NAN_METHOD(SetBorderBottom);
        static NAN_METHOD(BorderLeftColor);
        static NAN_METHOD(SetBorderLeftColor);
        static NAN_METHOD(BorderRightColor);
        static NAN_METHOD(SetBorderRightColor);
        static NAN_METHOD(BorderTopColor);
        static NAN_METHOD(SetBorderTopColor);
        static NAN_METHOD(BorderBottomColor);
        static NAN_METHOD(SetBorderBottomColor);
        static NAN_METHOD(BorderDiagonal);
        static NAN_METHOD(SetBorderDiagonal);
        static NAN_METHOD(BorderDiagonalColor);
        static NAN_METHOD(SetBorderDiagonalColor);
        static NAN_METHOD(FillPattern);
        static NAN_METHOD(SetFillPattern);
        static NAN_METHOD(PatternBackgroundColor);
        static NAN_METHOD(SetPatternBackgroundColor);
        static NAN_METHOD(PatternForegroundColor);
        static NAN_METHOD(SetPatternForegroundColor);
        static NAN_METHOD(Locked);
        static NAN_METHOD(SetLocked);
        static NAN_METHOD(Hidden);
        static NAN_METHOD(SetHidden);

    private:

        Format(const Format&);
        const Format& operator=(const Format&);
};


}

#endif // BINDINGS_FORMAT_H
