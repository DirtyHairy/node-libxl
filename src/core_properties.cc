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

#include "core_properties.h"

#include "argument_helper.h"
#include "assert.h"
#include "font.h"
#include "util.h"

using namespace v8;

namespace node_libxl {
    // Lifecycle

    CoreProperties::CoreProperties(libxl::CoreProperties* coreProperties, Local<Value> book)
        : Wrapper<libxl::CoreProperties>(coreProperties), BookWrapper(book) {}

    Local<Object> CoreProperties::NewInstance(libxl::CoreProperties* libxlCoreProperties,
                                              Local<Value> book) {
        Nan::EscapableHandleScope scope;

        CoreProperties* coreProperties = new CoreProperties(libxlCoreProperties, book);

        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();

        coreProperties->Wrap(that);

        return scope.Escape(that);
    }

    // Wrappers

    NAN_METHOD(CoreProperties::Title) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* title = that->GetWrapped()->title();
        if (!title) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(title).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetTitle) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(title, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setTitle(*title);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Subject) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* subject = that->GetWrapped()->subject();
        if (!subject) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(subject).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetSubject) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(subject, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setSubject(*subject);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Creator) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* creator = that->GetWrapped()->creator();
        if (!creator) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(creator).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetCreator) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(creator, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCreator(*creator);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::LastModifiedBy) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* lastModifiedBy = that->GetWrapped()->lastModifiedBy();
        if (!lastModifiedBy) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(lastModifiedBy).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetLastModifiedBy) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(lastModifiedBy, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setLastModifiedBy(*lastModifiedBy);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Created) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* created = that->GetWrapped()->created();
        if (!created) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(created).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetCreated) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(created, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCreated(*created);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::CreatedAsDouble) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        double created = that->GetWrapped()->createdAsDouble();
        info.GetReturnValue().Set(Nan::New<Number>(created));
    }

    NAN_METHOD(CoreProperties::SetCreatedAsDouble) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double created = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCreatedAsDouble(created);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Modified) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* modified = that->GetWrapped()->modified();
        if (!modified) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(modified).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetModified) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(modified, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setModified(*modified);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::ModifiedAsDouble) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        double modified = that->GetWrapped()->modifiedAsDouble();
        info.GetReturnValue().Set(Nan::New<Number>(modified));
    }

    NAN_METHOD(CoreProperties::SetModifiedAsDouble) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        double modified = arguments.GetDouble(0);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setModifiedAsDouble(modified);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Tags) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* tags = that->GetWrapped()->tags();
        if (!tags) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(tags).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetTags) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(tags, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setTags(*tags);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Categories) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* categories = that->GetWrapped()->categories();
        if (!categories) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(categories).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetCategories) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(categories, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setCategories(*categories);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::Comments) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        const char* comments = that->GetWrapped()->comments();
        if (!comments) {
            return util::ThrowLibxlError(that);
        }

        info.GetReturnValue().Set(Nan::New<String>(comments).ToLocalChecked());
    }

    NAN_METHOD(CoreProperties::SetComments) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        CSNanUtf8Value(comments, arguments.GetString(0));
        ASSERT_ARGUMENTS(arguments);

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->setComments(*comments);

        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(CoreProperties::RemoveAll) {
        Nan::HandleScope scope;

        CoreProperties* that = Unwrap(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->removeAll();

        info.GetReturnValue().Set(info.This());
    }

    // Init

    void CoreProperties::Initialize(Local<Object> exports) {
        using namespace libxl;

        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("CoreProperties").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookWrapper::Initialize<CoreProperties>(t);

        Nan::SetPrototypeMethod(t, "title", Title);
        Nan::SetPrototypeMethod(t, "setTitle", SetTitle);
        Nan::SetPrototypeMethod(t, "subject", Subject);
        Nan::SetPrototypeMethod(t, "setSubject", SetSubject);
        Nan::SetPrototypeMethod(t, "creator", Creator);
        Nan::SetPrototypeMethod(t, "setCreator", SetCreator);
        Nan::SetPrototypeMethod(t, "lastModifiedBy", LastModifiedBy);
        Nan::SetPrototypeMethod(t, "setLastModifiedBy", SetLastModifiedBy);
        Nan::SetPrototypeMethod(t, "created", Created);
        Nan::SetPrototypeMethod(t, "setCreated", SetCreated);
        Nan::SetPrototypeMethod(t, "createdAsDouble", CreatedAsDouble);
        Nan::SetPrototypeMethod(t, "setCreatedAsDouble", SetCreatedAsDouble);
        Nan::SetPrototypeMethod(t, "modified", Modified);
        Nan::SetPrototypeMethod(t, "setModified", SetModified);
        Nan::SetPrototypeMethod(t, "modifiedAsDouble", ModifiedAsDouble);
        Nan::SetPrototypeMethod(t, "setModifiedAsDouble", SetModifiedAsDouble);
        Nan::SetPrototypeMethod(t, "tags", Tags);
        Nan::SetPrototypeMethod(t, "setTags", SetTags);
        Nan::SetPrototypeMethod(t, "categories", Categories);
        Nan::SetPrototypeMethod(t, "setCategories", SetCategories);
        Nan::SetPrototypeMethod(t, "comments", Comments);
        Nan::SetPrototypeMethod(t, "setComments", SetComments);
        Nan::SetPrototypeMethod(t, "removeAll", RemoveAll);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("CoreProperties").ToLocalChecked(), Nan::New(constructor));
    }
}  // namespace node_libxl