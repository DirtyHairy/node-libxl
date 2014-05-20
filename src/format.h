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

        static v8::Handle<v8::Object> NewInstance(
            libxl::Format* format,
            v8::Handle<v8::Value> book
        );

    protected:

        static NAN_METHOD(SetFont);
        static NAN_METHOD(SetNumFormat);
        static NAN_METHOD(NumFormat);
        static NAN_METHOD(SetWrap);
        static NAN_METHOD(SetShrinkToFit);
        static NAN_METHOD(SetAlignH);
        static NAN_METHOD(SetFillPattern);
        static NAN_METHOD(SetPatternBackgroundColor);
        static NAN_METHOD(SetPatternForegroundColor);

    private:

        Format(const Format&);
        const Format& operator=(const Format&);
};


}

#endif // BINDINGS_FORMAT_H
