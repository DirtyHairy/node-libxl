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

#ifndef NODE_LIBXL_FORM_CONTROL_H
#define NODE_LIBXL_FORM_CONTROL_H

#include "book_holder.h"
#include "common.h"
#include "wrapper.h"

namespace node_libxl {
    class FormControl : public Wrapper<libxl::IFormControlT<char>, FormControl>, public BookHolder {
       public:
        FormControl(libxl::IFormControlT<char>* formControl, v8::Local<v8::Value> book);

        static void Initialize(v8::Local<v8::Object> exports);

        static v8::Local<v8::Object> NewInstance(libxl::IFormControlT<char>* libxlFormControl,
                                                 v8::Local<v8::Value> book);

       protected:
        static NAN_METHOD(ObjectType);
        static NAN_METHOD(Checked);
        static NAN_METHOD(SetChecked);
        static NAN_METHOD(FmlaGroup);
        static NAN_METHOD(SetFmlaGroup);
        static NAN_METHOD(FmlaLink);
        static NAN_METHOD(SetFmlaLink);
        static NAN_METHOD(FmlaRange);
        static NAN_METHOD(SetFmlaRange);
        static NAN_METHOD(FmlaTxbx);
        static NAN_METHOD(SetFmlaTxbx);
        static NAN_METHOD(Name);
        static NAN_METHOD(LinkedCell);
        static NAN_METHOD(ListFillRange);
        static NAN_METHOD(Macro);
        static NAN_METHOD(AltText);
        static NAN_METHOD(Locked);
        static NAN_METHOD(DefaultSize);
        static NAN_METHOD(Print);
        static NAN_METHOD(Disabled);
        static NAN_METHOD(Item);
        static NAN_METHOD(ItemSize);
        static NAN_METHOD(AddItem);
        static NAN_METHOD(InsertItem);
        static NAN_METHOD(ClearItems);
        static NAN_METHOD(DropLines);
        static NAN_METHOD(SetDropLines);
        static NAN_METHOD(Dx);
        static NAN_METHOD(SetDx);
        static NAN_METHOD(FirstButton);
        static NAN_METHOD(SetFirstButton);
        static NAN_METHOD(Horiz);
        static NAN_METHOD(SetHoriz);
        static NAN_METHOD(Inc);
        static NAN_METHOD(SetInc);
        static NAN_METHOD(GetMax);
        static NAN_METHOD(SetMax);
        static NAN_METHOD(GetMin);
        static NAN_METHOD(SetMin);
        static NAN_METHOD(MultiSel);
        static NAN_METHOD(SetMultiSel);
        static NAN_METHOD(Sel);
        static NAN_METHOD(SetSel);
        static NAN_METHOD(FromAnchor);
        static NAN_METHOD(ToAnchor);

       private:
        FormControl(const FormControl&);
        FormControl& operator=(const FormControl&);
    };
}  // namespace node_libxl

#endif  // NODE_LIBXL_FORM_CONTROL_H
