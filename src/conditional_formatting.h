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

#ifndef NODE_LIBXL_CONDITIONAL_FORMATTING_H
#define NODE_LIBXL_CONDITIONAL_FORMATTING_H

#include "book_holder.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {
    class ConditionalFormatting
        : public Wrapper<libxl::ConditionalFormatting, ConditionalFormatting>,
          public BookHolder {
       public:
        ConditionalFormatting(libxl::ConditionalFormatting* conditionalFormatting,
                              v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static v8::Local<v8::Object> NewInstance(
            libxl::ConditionalFormatting* libxlConditionalFormatting, v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(AddRange);
        static NAN_METHOD(AddRule);
        static NAN_METHOD(AddTopRule);
        static NAN_METHOD(AddOpNumRule);
        static NAN_METHOD(AddOpStrRule);
        static NAN_METHOD(AddAboveAverageRule);
        static NAN_METHOD(AddTimePeriodRule);
        static NAN_METHOD(Add2ColorScaleRule);
        static NAN_METHOD(Add2ColorScaleFormulaRule);
        static NAN_METHOD(Add3ColorScaleRule);
        static NAN_METHOD(Add3ColorScaleFormulaRule);

       private:
        ConditionalFormatting(const ConditionalFormatting&);
        ConditionalFormatting& operator=(const ConditionalFormatting&);
    };
}  // namespace node_libxl

#endif  // NODE_LIBXL_CONDITIONAL_FORMATTING_H
