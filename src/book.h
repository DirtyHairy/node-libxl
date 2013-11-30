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

#ifndef BINDINGS_BOOK
#define BINDINGS_BOOK

#define BUILDING_NODE_EXTENSION

#include <node.h>
#include <v8.h>
#include <libxl.h>

#include "wrapper.h"

namespace node_libxl {


enum {
    BOOK_TYPE_XLS,
    BOOK_TYPE_XLSX
};


class Book : public Wrapper<libxl::Book> {
    public:

        Book(libxl::Book* libxlBook);
        ~Book();

        static void Initialize(v8::Handle<v8::Object> exports);

        static Book* Unwrap(v8::Handle<v8::Value> object) {
            return Wrapper<libxl::Book>::Unwrap<Book>(object);
        }

    protected:

        static v8::Handle<v8::Value> New(const v8::Arguments& arguments);

        static v8::Handle<v8::Value> WriteSync(const v8::Arguments& arguments);
        static v8::Handle<v8::Value> AddSheet(const v8::Arguments& arguments);

    private:

        Book(const Book&);
        const Book& operator=(const Book&);

};


}

#include "libxl.h"

#endif // BINDINGS_BOOK
