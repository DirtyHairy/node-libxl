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

#ifndef BINDINGS_STRING_COPY_H
#define BINDINGS_STRING_COPY_H

#include <v8.h>

#include <optional>

namespace node_libxl {

    class StringCopy {
       public:
        explicit StringCopy(v8::String::Utf8Value& utf8Value);
        explicit StringCopy(v8::Local<v8::Value> value);
        explicit StringCopy(std::optional<v8::Local<v8::Value>> value);

        ~StringCopy();

        char* operator*();

       private:
        StringCopy(const StringCopy&) = delete;
        StringCopy(StringCopy&&) = delete;
        StringCopy& operator=(const StringCopy&);
        StringCopy& operator=(StringCopy&&);

        char* str{nullptr};
    };

}  // namespace node_libxl

#endif  // BINDINGS_STRING_COPY_H
