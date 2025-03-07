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

#ifndef BINDINGS_WRAPPER_H
#define BINDINGS_WRAPPER_H

#include "common.h"

namespace node_libxl {

    template <typename T, typename U>
    class Wrapper : public Nan::ObjectWrap {
       public:
        Wrapper() : wrapped(NULL) {}
        Wrapper(T* wrapped) : wrapped(wrapped) {}

        T* GetWrapped() { return wrapped; }

        static bool InstanceOf(v8::Local<v8::Value> object);
        static U* FromJS(v8::Local<v8::Value> object);

       protected:
        static Nan::Persistent<v8::Function> constructor;
        T* wrapped;

       private:
        Wrapper(const Wrapper<T, U>&);
        const Wrapper<T, U>& operator=(const Wrapper<T, U>&);
    };

    // Implementation

    template <typename T, typename U>
    Nan::Persistent<v8::Function> Wrapper<T, U>::constructor;

    template <typename T, typename U>
    bool Wrapper<T, U>::InstanceOf(v8::Local<v8::Value> object) {
        Nan::HandleScope scope;

        return object->IsObject() &&
               object.As<v8::Object>()->GetPrototype()->StrictEquals(
                   Nan::Get(Nan::New(constructor),
                            Nan::New<v8::String>("prototype").ToLocalChecked())
                       .ToLocalChecked());
    }

    template <typename T, typename U>
    U* Wrapper<T, U>::FromJS(v8::Local<v8::Value> object) {
        if (InstanceOf(object)) {
            return Nan::ObjectWrap::Unwrap<U>(object.As<v8::Object>());
        } else {
            return NULL;
        }
    }

}  // namespace node_libxl

#endif  // BINDINGS_WRAPPER_H
