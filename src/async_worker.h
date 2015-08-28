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

#ifndef BINDINGS_ASYNC_WORKER_H
#define BINDINGS_ASYNC_WORKER_H

#include <v8.h>
#include <nan.h>
#include "util.h"

namespace node_libxl {


template<typename T> class AsyncWorker : public Nan::AsyncWorker {
    public:

        AsyncWorker(Nan::Callback* callback, v8::Local<v8::Object> that);

        virtual void WorkComplete();

    protected:

        void RaiseLibxlError();

        T* that;

    private:

        AsyncWorker(const AsyncWorker&);
        const AsyncWorker& operator=(const AsyncWorker&);
};


template<typename T> AsyncWorker<T>::AsyncWorker(
        Nan::Callback* callback, v8::Local<v8::Object> that) :
    Nan::AsyncWorker(callback),
    that(T::Unwrap(that))
{
    util::GetBook(this->that)->StartAsync();
    SaveToPersistent("that", that);
}


template<typename T> void AsyncWorker<T>::WorkComplete() {
    util::GetBook(that)->StopAsync();

    Nan::AsyncWorker::WorkComplete();
}


template<typename T> void AsyncWorker<T>::RaiseLibxlError() {
    SetErrorMessage(util::UnwrapBook(that)->errorMessage());
}


}

#endif // BINDINGS_ASYNC_WORKER_H
