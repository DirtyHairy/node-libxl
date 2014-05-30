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

#include "argument_helper.h"

#include <sstream>

namespace node_libxl {


ArgumentHelper::ArgumentHelper(_NAN_METHOD_ARGS_TYPE args) :
    arguments(args),
    exceptionRaised(false)
{}


int32_t ArgumentHelper::GetInt(uint8_t pos) {
    if (!arguments[pos]->IsInt32()) {
        RaiseException("integer required at position", pos);
        return 0;
    }

    return arguments[pos]->Int32Value();
}


int32_t ArgumentHelper::GetInt(uint8_t pos, int32_t def) {
    if (arguments[pos]->IsUndefined()) return def;
    return GetInt(pos);
}


double ArgumentHelper::GetDouble(uint8_t pos) {
    if (!arguments[pos]->IsNumber()) {
        RaiseException("number required at position", pos);
        return 0;
    }

    return arguments[pos]->NumberValue();
}


double ArgumentHelper::GetDouble(uint8_t pos, double def) {
    if (arguments[pos]->IsUndefined()) return def;
    return GetDouble(pos);
}


bool ArgumentHelper::GetBoolean(uint8_t pos) {
    if (!arguments[pos]->IsBoolean()) {
        RaiseException("bool required at position", pos);
        return false;
    }

    return arguments[pos]->BooleanValue();
}


bool ArgumentHelper::GetBoolean(uint8_t pos, bool def) {
    if (arguments[pos]->IsUndefined()) return def;
    return GetBoolean(pos);
}


v8::Handle<v8::Value> ArgumentHelper::GetString(uint8_t pos) {
    if (!arguments[pos]->IsString()) {
        RaiseException("string required at position", pos);
    }

    return arguments[pos];
}


v8::Handle<v8::Value> ArgumentHelper::GetString(uint8_t pos, const char* def) {
    NanEscapableScope();

    if (arguments[pos]->IsUndefined())
        return NanEscapeScope(NanNew<v8::String>(def));

    return NanEscapeScope(GetString(pos));
}


void ArgumentHelper::RaiseException(const std::string& message, int32_t pos) {
    NanEscapableScope();

    if (HasException()) return;

    std::stringstream ss;
    ss.str("");
    ss << message;

    if (pos >= 0) ss << " " << pos;

    exceptionMessage = ss.str();
    exceptionRaised = true;
}


bool ArgumentHelper::HasException() const {
    return exceptionRaised;
}

_NAN_METHOD_RETURN_TYPE ArgumentHelper::ThrowException() const {
    return NanThrowTypeError(exceptionMessage.data());
}


}
