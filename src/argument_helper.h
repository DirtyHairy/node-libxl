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

#ifndef BINDINGS_ARGUMENT_HELPER_H
#define BINDINGS_ARGUMENT_HELPER_H

#include <optional>
#include <string>

#include "common.h"

namespace node_libxl {

    class ArgumentHelper {
       public:
        ArgumentHelper(Nan::NAN_METHOD_ARGS_TYPE info);

        size_t Length();

        int GetInt(size_t pos);
        std::optional<int> GetMaybeInt(size_t pos);
        int GetInt(size_t pos, int def);

        double GetDouble(size_t pos);
        std::optional<double> GetMaybeDouble(size_t pos);
        double GetDouble(size_t pos, double def);

        bool GetBoolean(size_t pos);
        std::optional<bool> GetMaybeBoolean(size_t pos);
        bool GetBoolean(size_t pos, bool def);

        v8::Local<v8::Value> GetString(size_t pos);
        std::optional<v8::Local<v8::Value>> GetMaybeString(size_t pos);
        v8::Local<v8::Value> GetString(size_t pos, const char *def);

        v8::Local<v8::Function> GetFunction(size_t pos);
        std::optional<v8::Local<v8::Function>> GetMaybeFunction(size_t pos);

        v8::Local<v8::Value> GetBuffer(size_t pos);
        std::optional<v8::Local<v8::Value>> GetMaybeBuffer(size_t pos);

        bool IsDefined(size_t pos);

        template <typename T>
        T *GetWrapped(size_t pos);
        template <typename T>
        T *GetWrapped(size_t pos, T *def);

        bool HasException() const;
        Nan::NAN_METHOD_RETURN_TYPE ThrowException() const;

       private:
        Nan::NAN_METHOD_ARGS_TYPE arguments;
        std::string exceptionMessage;
        bool exceptionRaised;

        void RaiseException(const std::string &message, int32_t pos = -1);

        ArgumentHelper(const ArgumentHelper &);
        const ArgumentHelper &operator=(ArgumentHelper &);
    };

    template <typename T>
    T *ArgumentHelper::GetWrapped(size_t pos) {
        T *unwrapped = T::FromJS(arguments[pos]);
        if (!unwrapped) RaiseException("Invalid type for argument", pos);
        return unwrapped;
    }

    template <typename T>
    T *ArgumentHelper::GetWrapped(size_t pos, T *def) {
        if (arguments[pos]->IsUndefined()) return def;
        return GetWrapped<T>(pos);
    }

}  // namespace node_libxl

#endif  // BINDINGS_ARGUMENT_HELPER_H
