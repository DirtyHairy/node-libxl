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


template<typename T> class Wrapper : public node::ObjectWrap {
    public:

        Wrapper() : wrapped(NULL) {}
        Wrapper(T* wrapped) : wrapped(wrapped) {}

        T* GetWrapped() {
            return wrapped;
        }

        static v8::Handle<v8::Function> GetConstructor() {
            NanScope();

            return NanNew(constructor);
        }

        static bool InstanceOf(v8::Handle<v8::Value> object);
        template<typename U> static U* Unwrap(v8::Handle<v8::Value> object);

    protected:

        static v8::Persistent<v8::Function> constructor;
        T* wrapped;

    private:

        Wrapper(const Wrapper<T>&);
        const Wrapper<T>& operator=(const Wrapper<T>&);
};


// Implementation


template<typename T> v8::Persistent<v8::Function> Wrapper<T>::constructor;


template<typename T> bool Wrapper<T>::InstanceOf(v8::Handle<v8::Value> object) {
    NanScope();

    return object->IsObject() &&
        object.As<v8::Object>()->GetPrototype()->StrictEquals(
            NanNew(constructor)->Get(NanSymbol("prototype")));
}


template <typename T> template<typename U>
    U* Wrapper<T>::Unwrap(v8::Handle<v8::Value> object)
{
    if (InstanceOf(object)) {
        return node::ObjectWrap::Unwrap<U>(object.As<v8::Object>());
    } else {
        return NULL;
    }
}


}

#endif // BINDINGS_WRAPPER_H
