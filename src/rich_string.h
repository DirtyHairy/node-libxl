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

#ifndef NODE_LIBXL_RICH_STRING_H
#define NODE_LIBXL_RICH_STRING_H

#include "book_wrapper.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {
    class RichString : public Wrapper<libxl::RichString>, public BookWrapper {
       public:
        RichString(libxl::RichString* richString, v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static RichString* Unwrap(v8::Local<v8::Value> object) {
            return Wrapper<libxl::RichString>::Unwrap<RichString>(object);
        }

        static v8::Local<v8::Object> NewInstance(libxl::RichString* libxlRichString,
                                                 v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(AddFont);
        static NAN_METHOD(AddText);
        static NAN_METHOD(GetText);
        static NAN_METHOD(TextSize);

       private:
        RichString(const RichString&);
        RichString& operator=(const RichString&);
    };
}  // namespace node_libxl

#endif  // NODE_LIBXL_RICH_STRING_H
