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

    int ArgumentHelper::GetInt(size_t pos) {
        if (!arguments[pos]->IsInt32()) {
            RaiseException("integer required at position", pos);
            return 0;
        }

        return arguments[pos]->IntegerValue(Nan::GetCurrentContext()).ToChecked();
    }

    size_t ArgumentHelper::Length() { return arguments.Length(); }

    int ArgumentHelper::GetInt(size_t pos, int def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetInt(pos);
    }

    std::optional<int> ArgumentHelper::GetMaybeInt(size_t pos) {
        return IsDefined(pos) ? std::optional(GetInt(pos)) : std::nullopt;
    }

    double ArgumentHelper::GetDouble(size_t pos) {
        if (!arguments[pos]->IsNumber()) {
            RaiseException("number required at position", pos);
            return 0;
        }

        return arguments[pos]->NumberValue(Nan::GetCurrentContext()).ToChecked();
    }

    double ArgumentHelper::GetDouble(size_t pos, double def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetDouble(pos);
    }

    std::optional<double> ArgumentHelper::GetMaybeDouble(size_t pos) {
        return IsDefined(pos) ? std::optional(GetDouble(pos)) : std::nullopt;
    }

    bool ArgumentHelper::GetBoolean(size_t pos) {
        if (!arguments[pos]->IsBoolean()) {
            RaiseException("bool required at position", pos);
            return false;
        }

        return arguments[pos]->BooleanValue(v8::Isolate::GetCurrent());
    }

    bool ArgumentHelper::IsDefined(size_t pos) { return !arguments[pos]->IsUndefined(); }

    bool ArgumentHelper::GetBoolean(size_t pos, bool def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetBoolean(pos);
    }

    std::optional<bool> ArgumentHelper::GetMaybeBoolean(size_t pos) {
        return IsDefined(pos) ? std::optional(GetBoolean(pos)) : std::nullopt;
    }

    v8::Local<v8::Value> ArgumentHelper::GetString(size_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsString()) {
            RaiseException("string required at position", pos);
        }

        return scope.Escape(arguments[pos]);
    }

    v8::Local<v8::Value> ArgumentHelper::GetString(size_t pos, const char *def) {
        Nan::EscapableHandleScope scope;

        if (arguments[pos]->IsUndefined())
            return scope.Escape(Nan::New<v8::String>(def).ToLocalChecked());

        return scope.Escape(GetString(pos));
    }

    std::optional<v8::Local<v8::Value>> ArgumentHelper::GetMaybeString(size_t pos) {
        return IsDefined(pos) ? std::optional(GetString(pos)) : std::nullopt;
    }

    v8::Local<v8::Function> ArgumentHelper::GetFunction(size_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsFunction()) {
            RaiseException("function required at position", pos);
        }

        return scope.Escape(arguments[pos].As<v8::Function>());
    }

    std::optional<v8::Local<v8::Function>> ArgumentHelper::GetMaybeFunction(size_t pos) {
        return IsDefined(pos) ? std::optional(GetFunction(pos)) : std::nullopt;
    }

    v8::Local<v8::Value> ArgumentHelper::GetBuffer(size_t pos) {
        Nan::EscapableHandleScope scope;

        if (!arguments[pos]->IsObject() || !node::Buffer::HasInstance(arguments[pos])) {
            RaiseException("buffer required at position", pos);
        }

        return scope.Escape(arguments[pos]);
    }

    std::optional<v8::Local<v8::Value>> ArgumentHelper::GetMaybeBuffer(size_t pos) {
        return IsDefined(pos) ? std::optional(GetBuffer(pos)) : std::nullopt;
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
