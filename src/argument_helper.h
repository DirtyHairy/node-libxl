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

#include <string>

#include "common.h"

namespace node_libxl {


class ArgumentHelper {
    public:

        ArgumentHelper(_NAN_METHOD_ARGS_TYPE args);

        int32_t GetInt(uint8_t pos);
        int32_t GetInt(uint8_t pos, int32_t def);

        double GetDouble(uint8_t pos);
        double GetDouble(uint8_t pos, double def);

        bool GetBoolean(uint8_t pos);
        bool GetBoolean(uint8_t pos, bool def);

        v8::Handle<v8::Value> GetString(uint8_t pos);
        v8::Handle<v8::Value> GetString(uint8_t pos, const char* def);

        bool HasException() const;
        _NAN_METHOD_RETURN_TYPE ThrowException() const;

    private:

        _NAN_METHOD_ARGS_TYPE arguments;
        std::string exceptionMessage;
        bool exceptionRaised;

        void RaiseException(const std::string& message, int32_t pos = -1);

        ArgumentHelper(const ArgumentHelper&);
        const ArgumentHelper& operator=(ArgumentHelper&);
};


}

#endif // BINDINGS_ARGUMENT_HELPER_H
