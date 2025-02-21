/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2025 Christian Speckner <cnspeckn@googlemail.com>
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

#ifndef BINDINGS_CORE_PROPERTIES_H
#define BINDINGS_CORE_PROPERTIES_H

#include "book_holder.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {

    class CoreProperties : public Wrapper<libxl::CoreProperties, CoreProperties>,
                           public BookHolder

    {
       public:
        CoreProperties(libxl::CoreProperties* coreProperties, v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static v8::Local<v8::Object> NewInstance(libxl::CoreProperties* coreProperties,
                                                 v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(Title);
        static NAN_METHOD(SetTitle);
        static NAN_METHOD(Subject);
        static NAN_METHOD(SetSubject);
        static NAN_METHOD(Creator);
        static NAN_METHOD(SetCreator);
        static NAN_METHOD(LastModifiedBy);
        static NAN_METHOD(SetLastModifiedBy);
        static NAN_METHOD(Created);
        static NAN_METHOD(SetCreated);
        static NAN_METHOD(CreatedAsDouble);
        static NAN_METHOD(SetCreatedAsDouble);
        static NAN_METHOD(Modified);
        static NAN_METHOD(ModifiedAsDouble);
        static NAN_METHOD(SetModifiedAsDouble);
        static NAN_METHOD(Tags);
        static NAN_METHOD(SetTags);
        static NAN_METHOD(Categories);
        static NAN_METHOD(SetCategories);
        static NAN_METHOD(Comments);
        static NAN_METHOD(SetComments);
        static NAN_METHOD(RemoveAll);
        static NAN_METHOD(SetModified);

       private:
        CoreProperties(const CoreProperties&);
        const CoreProperties& operator=(const CoreProperties&);
    };

}  // namespace node_libxl

#endif  // BINDINGS_CORE_PROPERTIES_H
