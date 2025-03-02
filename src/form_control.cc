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

#include "form_control.h"

#include "argument_helper.h"
#include "assert.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    FormControl::FormControl(libxl::IFormControlT<char>* formControl, Local<Value> book)
        : Wrapper<libxl::IFormControlT<char>, FormControl>(formControl), BookHolder(book) {}

    Local<Object> FormControl::NewInstance(libxl::IFormControlT<char>* libxlFormControl,
                                           Local<Value> book) {
        Nan::EscapableHandleScope scope;

        FormControl* formControl = new FormControl(libxlFormControl, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        formControl->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(FormControl::ObjectType) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->objectType()));
    }

    NAN_METHOD(FormControl::Checked) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->checked()));
    }

    NAN_METHOD(FormControl::SetChecked) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int checked = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setChecked(static_cast<libxl::CheckedType>(checked));
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FmlaGroup) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(
            Nan::New<String>(that->GetWrapped()->fmlaGroup()).ToLocalChecked());
    }

    NAN_METHOD(FormControl::SetFmlaGroup) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(group, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFmlaGroup(*group);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FmlaLink) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->fmlaLink();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::SetFmlaLink) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(link, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFmlaLink(*link);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FmlaRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->fmlaRange();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::SetFmlaRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(range, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());

        ASSERT_THIS(that);
        that->GetWrapped()->setFmlaRange(*range);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FmlaTxbx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->fmlaTxbx();

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::SetFmlaTxbx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(txbx, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFmlaTxbx(*txbx);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::Name) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->name();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::LinkedCell) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->linkedCell();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::ListFillRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->listFillRange();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::Macro) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->macro();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::AltText) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->altText();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::Locked) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->locked()));
    }

    NAN_METHOD(FormControl::DefaultSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->defaultSize()));
    }

    NAN_METHOD(FormControl::Print) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->print()));
    }

    NAN_METHOD(FormControl::Disabled) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->disabled()));
    }

    NAN_METHOD(FormControl::Item) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->item(index);
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::ItemSize) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->itemSize()));
    }

    NAN_METHOD(FormControl::AddItem) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(value, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->addItem(*value);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::InsertItem) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int index = arguments.GetInt(0);
        CSNanUtf8Value(value, arguments.GetString(1));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->insertItem(index, *value);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::ClearItems) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->clearItems();
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::DropLines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->dropLines()));
    }

    NAN_METHOD(FormControl::SetDropLines) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int lines = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setDropLines(lines);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::Dx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->dx()));
    }

    NAN_METHOD(FormControl::SetDx) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int dx = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setDx(dx);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FirstButton) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->firstButton()));
    }

    NAN_METHOD(FormControl::SetFirstButton) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool firstButton = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setFirstButton(firstButton);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::Horiz) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Boolean>(that->GetWrapped()->horiz()));
    }

    NAN_METHOD(FormControl::SetHoriz) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        bool horiz = arguments.GetBoolean(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setHoriz(horiz);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::Inc) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->inc()));
    }

    NAN_METHOD(FormControl::SetInc) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int inc = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setInc(inc);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::GetMax) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->getMax()));
    }

    NAN_METHOD(FormControl::SetMax) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int max = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setMax(max);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::GetMin) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->getMin()));
    }

    NAN_METHOD(FormControl::SetMin) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int min = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setMin(min);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::MultiSel) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        const char* result = that->GetWrapped()->multiSel();
        if (!result) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(result).ToLocalChecked());
    }

    NAN_METHOD(FormControl::SetMultiSel) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(value, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setMultiSel(*value);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::Sel) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        info.GetReturnValue().Set(Nan::New<Number>(that->GetWrapped()->sel()));
    }

    NAN_METHOD(FormControl::SetSel) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int sel = arguments.GetInt(0);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setSel(sel);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(FormControl::FromAnchor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        int col, colOff, row, rowOff;
        if (!that->GetWrapped()->fromAnchor(&col, &colOff, &row, &rowOff)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Number>(col));
        Nan::Set(result, Nan::New<String>("colOff").ToLocalChecked(), Nan::New<Number>(colOff));
        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Number>(row));
        Nan::Set(result, Nan::New<String>("rowOff").ToLocalChecked(), Nan::New<Number>(rowOff));

        info.GetReturnValue().Set(result);
    }

    NAN_METHOD(FormControl::ToAnchor) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        FormControl* that = FromJS(info.This());
        ASSERT_THIS(that);

        int col, colOff, row, rowOff;
        if (!that->GetWrapped()->toAnchor(&col, &colOff, &row, &rowOff)) {
            return util::ThrowLibxlError(that);
        }

        Local<Object> result = Nan::New<Object>();
        Nan::Set(result, Nan::New<String>("col").ToLocalChecked(), Nan::New<Number>(col));
        Nan::Set(result, Nan::New<String>("colOff").ToLocalChecked(), Nan::New<Number>(colOff));
        Nan::Set(result, Nan::New<String>("row").ToLocalChecked(), Nan::New<Number>(row));
        Nan::Set(result, Nan::New<String>("rowOff").ToLocalChecked(), Nan::New<Number>(rowOff));

        info.GetReturnValue().Set(result);
    }

    // Init

    void FormControl::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("FormControl").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<FormControl>(t);

        Nan::SetPrototypeMethod(t, "objectType", ObjectType);
        Nan::SetPrototypeMethod(t, "checked", Checked);
        Nan::SetPrototypeMethod(t, "setChecked", SetChecked);
        Nan::SetPrototypeMethod(t, "fmlaGroup", FmlaGroup);
        Nan::SetPrototypeMethod(t, "setFmlaGroup", SetFmlaGroup);
        Nan::SetPrototypeMethod(t, "fmlaLink", FmlaLink);
        Nan::SetPrototypeMethod(t, "setFmlaLink", SetFmlaLink);
        Nan::SetPrototypeMethod(t, "fmlaRange", FmlaRange);
        Nan::SetPrototypeMethod(t, "setFmlaRange", SetFmlaRange);
        Nan::SetPrototypeMethod(t, "fmlaTxbx", FmlaTxbx);
        Nan::SetPrototypeMethod(t, "setFmlaTxbx", SetFmlaTxbx);
        Nan::SetPrototypeMethod(t, "name", Name);
        Nan::SetPrototypeMethod(t, "linkedCell", LinkedCell);
        Nan::SetPrototypeMethod(t, "listFillRange", ListFillRange);
        Nan::SetPrototypeMethod(t, "macro", Macro);
        Nan::SetPrototypeMethod(t, "altText", AltText);
        Nan::SetPrototypeMethod(t, "locked", Locked);
        Nan::SetPrototypeMethod(t, "defaultSize", DefaultSize);
        Nan::SetPrototypeMethod(t, "print", Print);
        Nan::SetPrototypeMethod(t, "disabled", Disabled);
        Nan::SetPrototypeMethod(t, "item", Item);
        Nan::SetPrototypeMethod(t, "itemSize", ItemSize);
        Nan::SetPrototypeMethod(t, "addItem", AddItem);
        Nan::SetPrototypeMethod(t, "insertItem", InsertItem);
        Nan::SetPrototypeMethod(t, "clearItems", ClearItems);
        Nan::SetPrototypeMethod(t, "dropLines", DropLines);
        Nan::SetPrototypeMethod(t, "setDropLines", SetDropLines);
        Nan::SetPrototypeMethod(t, "dx", Dx);
        Nan::SetPrototypeMethod(t, "setDx", SetDx);
        Nan::SetPrototypeMethod(t, "firstButton", FirstButton);
        Nan::SetPrototypeMethod(t, "setFirstButton", SetFirstButton);
        Nan::SetPrototypeMethod(t, "horiz", Horiz);
        Nan::SetPrototypeMethod(t, "setHoriz", SetHoriz);
        Nan::SetPrototypeMethod(t, "inc", Inc);
        Nan::SetPrototypeMethod(t, "setInc", SetInc);
        Nan::SetPrototypeMethod(t, "getMax", GetMax);
        Nan::SetPrototypeMethod(t, "setMax", SetMax);
        Nan::SetPrototypeMethod(t, "getMin", GetMin);
        Nan::SetPrototypeMethod(t, "setMin", SetMin);
        Nan::SetPrototypeMethod(t, "multiSel", MultiSel);
        Nan::SetPrototypeMethod(t, "setMultiSel", SetMultiSel);
        Nan::SetPrototypeMethod(t, "sel", Sel);
        Nan::SetPrototypeMethod(t, "setSel", SetSel);
        Nan::SetPrototypeMethod(t, "fromAnchor", FromAnchor);
        Nan::SetPrototypeMethod(t, "toAnchor", ToAnchor);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("FormControl").ToLocalChecked(), Nan::New(constructor));
    }
}  // namespace node_libxl
