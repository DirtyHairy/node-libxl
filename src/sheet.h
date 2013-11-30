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

#include <v8.h>
#include <node.h>
#include <libxl.h>

#include "wrapper.h"

namespace node_libxl {


class Sheet : public Wrapper<libxl::Sheet> {
    public:

        Sheet(libxl::Sheet* sheet, v8::Handle<v8::Value> book);
        ~Sheet();

        static void Initialize(v8::Handle<v8::Object> exports);
        
        static Sheet* Unwrap(v8::Handle<v8::Value> object) {
            return Wrapper<libxl::Sheet>::Unwrap<Sheet>(object);
        }

        static v8::Handle<v8::Object> NewInstance(
            libxl::Sheet* sheet,
            v8::Handle<v8::Value> book
        );

        v8::Handle<v8::Value> GetBookHandle() const {
            return bookHandle;
        }

    protected:

        v8::Persistent<v8::Value> bookHandle;

        static v8::Handle<v8::Value> WriteString(const v8::Arguments&);
        static v8::Handle<v8::Value> WriteNum(const v8::Arguments&);
        static v8::Handle<v8::Value> WriteFormula(const v8::Arguments&);

    private:

        Sheet(const Sheet&);
        const Sheet& operator=(const Sheet&);
};


}
