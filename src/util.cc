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

#include "util.h"

#include <libxl.h>
#include <nan.h>

#include "book.h"

using namespace v8;


namespace node_libxl {
namespace util {

Handle<Value> ProxyConstructor(
    Handle<Function> constructor,
    _NAN_METHOD_ARGS_TYPE arguments
) {
    NanEscapableScope();

    uint32_t argc = arguments.Length();
    Handle<Value>* argv = new Handle<Value>[argc];

    for (uint32_t i = 0; i < argc; i++) {
        argv[i] = arguments[i];
    }

    Local<Value> newInstance = constructor->NewInstance(argc, argv);
    
    delete[] argv;
    return NanEscapeScope(newInstance);
}


NAN_METHOD(StubConstructor) {
    NanScope();

    Handle<Value> sentry = args[0];

    if (!(  args.IsConstructCall() &&
            args.Length() == 1 &&
            sentry->IsExternal() &&
            sentry.As<External>()->Value() == NULL
    )) {
        return NanThrowTypeError(
            "You are not supposed to call this constructor directly");
    }

    NanSetInternalFieldPointer(args.This(), 0, NULL);

    NanReturnValue(args.This());
}


Handle<Value> CallStubConstructor(Handle<Function> constructor) {
    NanEscapableScope();

    Handle<Value> args[1] = {CSNanNewExternal(NULL)};

    return NanEscapeScope(constructor->NewInstance(1, args));
}



}
}
