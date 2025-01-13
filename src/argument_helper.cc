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

    ArgumentHelper::ArgumentHelper(Nan::NAN_METHOD_ARGS_TYPE info)
        : arguments(info), exceptionRaised(false) {}

    int ArgumentHelper::GetInt(uint8_t pos) {
        if (!arguments[pos]->IsInt32()) {
            RaiseException("integer required at position", pos);
            return 0;
        }

        return arguments[pos]->IntegerValue(Nan::GetCurrentContext()).ToChecked();
    }

    int ArgumentHelper::GetInt(uint8_t pos, int def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetInt(pos);
    }

    double ArgumentHelper::GetDouble(uint8_t pos) {
        if (!arguments[pos]->IsNumber()) {
            RaiseException("number required at position", pos);
            return 0;
        }

        return arguments[pos]->NumberValue(Nan::GetCurrentContext()).ToChecked();
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

#if NODE_MAJOR_VERSION > 11
        return arguments[pos]->BooleanValue(v8::Isolate::GetCurrent());
#else
        return arguments[pos]->BooleanValue(Nan::GetCurrentContext()).ToChecked();
#endif
    }

    bool ArgumentHelper::GetBoolean(uint8_t pos, bool def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetBoolean(pos);
    }

    v8::Local<v8::Value> ArgumentHelper::GetString(uint8_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsString()) {
            RaiseException("string required at position", pos);
        }

        return scope.Escape(arguments[pos]);
    }

    v8::Local<v8::Value> ArgumentHelper::GetString(uint8_t pos, const char *def) {
        Nan::EscapableHandleScope scope;

        if (arguments[pos]->IsUndefined())
            return scope.Escape(Nan::New<v8::String>(def).ToLocalChecked());

        return scope.Escape(GetString(pos));
    }

    v8::Local<v8::Function> ArgumentHelper::GetFunction(uint8_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsFunction()) {
            RaiseException("function required at position", pos);
        }

        return scope.Escape(arguments[pos].As<v8::Function>());
    }

    v8::Local<v8::Value> ArgumentHelper::GetBuffer(uint8_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsObject() || !node::Buffer::HasInstance(arguments[pos])) {
            RaiseException("buffer required at position", pos);
        }

        return scope.Escape(arguments[pos]);
    }

    void ArgumentHelper::RaiseException(const std::string &message, int32_t pos) {
        Nan::EscapableHandleScope scope;

        if (HasException()) return;

        std::stringstream ss;
        ss.str("");
        ss << message;

        if (pos >= 0) ss << " " << pos;

        exceptionMessage = ss.str();
        exceptionRaised = true;
    }

    bool ArgumentHelper::HasException() const { return exceptionRaised; }

    Nan::NAN_METHOD_RETURN_TYPE ArgumentHelper::ThrowException() const {
        return Nan::ThrowTypeError(exceptionMessage.data());
    }

}  // namespace node_libxl
