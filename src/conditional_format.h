/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2025 Christian Speckner <cnspeckn@googlemail.com>
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

#ifndef NODE_LIBXL_CONDITIONAL_FORMAT_H
#define NODE_LIBXL_CONDITIONAL_FORMAT_H

#include "book_holder.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {
    class ConditionalFormat : public Wrapper<libxl::ConditionalFormat, ConditionalFormat>,
                              public BookHolder {
       public:
        ConditionalFormat(libxl::ConditionalFormat* conditionalFormat, v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static v8::Local<v8::Object> NewInstance(libxl::ConditionalFormat* libxlConditionalFormat,
                                                 v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(Font);
        static NAN_METHOD(NumFormat);
        static NAN_METHOD(SetNumFormat);
        static NAN_METHOD(CustomNumFormat);
        static NAN_METHOD(SetCustomNumFormat);
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
        static NAN_METHOD(FillPattern);
        static NAN_METHOD(SetFillPattern);
        static NAN_METHOD(PatternForegroundColor);
        static NAN_METHOD(SetPatternForegroundColor);
        static NAN_METHOD(PatternBackgroundColor);
        static NAN_METHOD(SetPatternBackgroundColor);

       private:
        ConditionalFormat(const ConditionalFormat&);
        ConditionalFormat& operator=(const ConditionalFormat&);
    };
}  // namespace node_libxl

#endif  // NODE_LIBXL_CONDITIONAL_FORMAT_H
