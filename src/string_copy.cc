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

#include "string_copy.h"
#include "csnan.h"
#include <cstring>

using namespace v8;

namespace node_libxl {

StringCopy::StringCopy(String::Utf8Value& utf8Value) {
    str = new char[strlen(*utf8Value) + 1];
    strcpy(str, *utf8Value);
}


StringCopy::StringCopy(Local<Value> value) {
    String::Utf8Value utf8Value(v8::Isolate::GetCurrent(), value);

    str = new char[strlen(*utf8Value) + 1];
    strcpy(str, *utf8Value);
}


StringCopy::~StringCopy() {
    delete[] str;
}


char* StringCopy::operator*() {
    return str;
}


}
