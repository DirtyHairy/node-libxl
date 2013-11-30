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

#include "book.h"

using namespace v8;


namespace node_libxl {
namespace util {

Handle<Value> ProxyConstructor(
    Handle<Function> constructor,
    const Arguments& arguments
) {
    HandleScope scope;

    uint32_t argc = arguments.Length();
    Handle<Value>* argv = new Handle<Value>[argc];

    for (uint32_t i = 0; i < argc; i++) {
        argv[i] = arguments[i];
    }

    Handle<Value> newInstance = constructor->NewInstance(argc, argv);
    
    delete[] argv;
    return scope.Close(newInstance);
}


Handle<Value> StubConstructor(const Arguments& arguments) {
    Handle<Value> sentry = arguments[0];
    if (!(  arguments.IsConstructCall() &&
            arguments.Length() == 1 &&
            sentry->IsExternal() &&
            sentry.As<External>()->Value() == NULL
    )) {
        return ThrowException(Exception::TypeError(String::New(
            "You are not supposed to call this constructor directly")));
    }

    arguments.This()->SetPointerInInternalField(0, NULL);

    return arguments.This();
}


Handle<Value> CallStubConstructor(Handle<Function> constructor) {
    HandleScope scope;

    Handle<Value> args[1] = {External::New(NULL)};
    return constructor->NewInstance(1, args);
}



}
}
