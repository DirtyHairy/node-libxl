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

#ifndef BINDINGS_CSNAN_H
#define BINDINGS_CSNAN_H

#define CSNanNewExternal(Value) v8::External::New(v8::Isolate::GetCurrent(), \
    Value)

#define CSNanObjectSetWithAttributes(Object,Key,Value,Attribs) \
    do {                                                                            \
        v8::Isolate* isolate = (Object)->GetIsolate();                              \
        v8::Local<v8::Context> context = isolate->GetCurrentContext();              \
        (Object)->DefineOwnProperty(context, Key, Value, Attribs).FromJust();       \
  } while (0)

#if NODE_MAJOR_VERSION >= 6
    #define CSNanNewInstance(handle, argv, argc) \
        (handle)->NewInstance((handle)->GetIsolate()->GetCurrentContext(), argv, argc).ToLocalChecked()
#else
    #define CSNanNewInstance(handle, argv, argc) (handle)->NewInstance(argv, argc)
#endif

#if NODE_MAJOR_VERSION >= 10
    #define CSNanUtf8Value(name, value) String::Utf8Value name(v8::Isolate::GetCurrent(), value)
#else
    #define CSNanUtf8Value(name, value) String::Utf8Value name(value)
#endif

#endif //BINDINGS_CSNAN_H
