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

#include "conditional_formatting.h"

#include "argument_helper.h"
#include "assert.h"
#include "conditional_format.h"
#include "util.h"

using namespace v8;

namespace node_libxl {

    ConditionalFormatting::ConditionalFormatting(
        libxl::ConditionalFormatting* conditionalFormatting, Local<Value> book)
        : Wrapper<libxl::ConditionalFormatting, ConditionalFormatting>(conditionalFormatting),
          BookHolder(book) {}

    Local<Object> ConditionalFormatting::NewInstance(
        libxl::ConditionalFormatting* libxlConditionalFormatting, Local<Value> book) {
        Nan::EscapableHandleScope scope;

        ConditionalFormatting* conditionalFormatting =
            new ConditionalFormatting(libxlConditionalFormatting, book);
        Local<Object> that = util::CallStubConstructor(Nan::New(constructor)).As<Object>();
        conditionalFormatting->Wrap(that);

        return scope.Escape(that);
    }

    NAN_METHOD(ConditionalFormatting::AddRange) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int rowFirst = arguments.GetInt(0);
        int rowLast = arguments.GetInt(1);
        int colFirst = arguments.GetInt(2);
        int colLast = arguments.GetInt(3);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->addRange(rowFirst, rowLast, colFirst, colLast);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int type = arguments.GetInt(0);
        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(1);
        CSNanUtf8Value(value, arguments.GetString(2, ""));
        bool stopIfTrue = arguments.GetBoolean(3, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addRule(static_cast<libxl::CFormatType>(type), cFormat->GetWrapped(),
                                    info[2]->IsUndefined() ? nullptr : *value, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddTopRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(0);
        int value = arguments.GetInt(1);
        bool bottom = arguments.GetBoolean(2, false);
        bool percent = arguments.GetBoolean(3, false);
        bool stopIfTrue = arguments.GetBoolean(4, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addTopRule(cFormat->GetWrapped(), value, bottom, percent, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddOpNumRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int op = arguments.GetInt(0);
        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(1);
        double value1 = arguments.GetDouble(2);
        double value2 = arguments.GetDouble(3, 0);
        bool stopIfTrue = arguments.GetBoolean(4, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addOpNumRule(static_cast<libxl::CFormatOperator>(op),
                                         cFormat->GetWrapped(), value1, value2, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddOpStrRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        int op = arguments.GetInt(0);
        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(1);
        CSNanUtf8Value(value1, arguments.GetString(2));
        CSNanUtf8Value(value2, arguments.GetString(3, ""));
        bool stopIfTrue = arguments.GetBoolean(4, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addOpStrRule(static_cast<libxl::CFormatOperator>(op),
                                         cFormat->GetWrapped(), *value1,
                                         info[3]->IsUndefined() ? nullptr : *value2, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddAboveAverageRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(0);
        bool aboveAverage = arguments.GetBoolean(1, true);
        bool equalAverage = arguments.GetBoolean(2, false);
        int stdDev = arguments.GetInt(3, 0);
        bool stopIfTrue = arguments.GetBoolean(4, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addAboveAverageRule(cFormat->GetWrapped(), aboveAverage, equalAverage,
                                                stdDev, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::AddTimePeriodRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);

        ConditionalFormat* cFormat = arguments.GetWrapped<ConditionalFormat>(0);
        int timePeriod = arguments.GetInt(1);
        bool stopIfTrue = arguments.GetBoolean(2, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);
        ASSERT_SAME_BOOK(cFormat, that);

        that->GetWrapped()->addTimePeriodRule(
            cFormat->GetWrapped(), static_cast<libxl::CFormatTimePeriod>(timePeriod), stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::Add2ColorScaleRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int minColor = arguments.GetInt(0);
        int maxColor = arguments.GetInt(1);
        int minType = arguments.GetInt(2, libxl::CFVO_MIN);
        double minValue = arguments.GetDouble(3, 0);
        int maxType = arguments.GetInt(4, libxl::CFVO_MAX);
        double maxValue = arguments.GetDouble(5, 0);
        bool stopIfTrue = arguments.GetBoolean(6, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->add2ColorScaleRule(
            static_cast<libxl::Color>(minColor), static_cast<libxl::Color>(maxColor),
            static_cast<libxl::CFVOType>(minType), minValue, static_cast<libxl::CFVOType>(maxType),
            maxValue, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::Add2ColorScaleFormulaRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int minColor = arguments.GetInt(0);
        int maxColor = arguments.GetInt(1);
        int minType = arguments.GetInt(2, libxl::CFVO_FORMULA);
        CSNanUtf8Value(minValue, arguments.GetString(3, ""));
        int maxType = arguments.GetInt(4, libxl::CFVO_FORMULA);
        CSNanUtf8Value(maxValue, arguments.GetString(5, ""));
        bool stopIfTrue = arguments.GetBoolean(6, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->add2ColorScaleFormulaRule(
            static_cast<libxl::Color>(minColor), static_cast<libxl::Color>(maxColor),
            static_cast<libxl::CFVOType>(minType), info[3]->IsUndefined() ? nullptr : *minValue,
            static_cast<libxl::CFVOType>(maxType), info[5]->IsUndefined() ? nullptr : *maxValue,
            stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::Add3ColorScaleRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int minColor = arguments.GetInt(0);
        int midColor = arguments.GetInt(1);
        int maxColor = arguments.GetInt(2);
        int minType = arguments.GetInt(3, libxl::CFVO_MIN);
        double minValue = arguments.GetDouble(4, 0);
        int midType = arguments.GetInt(5, libxl::CFVO_PERCENTILE);
        double midValue = arguments.GetDouble(6, 50);
        int maxType = arguments.GetInt(7, libxl::CFVO_MAX);
        double maxValue = arguments.GetDouble(8, 0);
        bool stopIfTrue = arguments.GetBoolean(9, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->add3ColorScaleRule(
            static_cast<libxl::Color>(minColor), static_cast<libxl::Color>(midColor),
            static_cast<libxl::Color>(maxColor), static_cast<libxl::CFVOType>(minType), minValue,
            static_cast<libxl::CFVOType>(midType), midValue, static_cast<libxl::CFVOType>(maxType),
            maxValue, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    NAN_METHOD(ConditionalFormatting::Add3ColorScaleFormulaRule) {
        Nan::HandleScope scope;

        ArgumentHelper arguments(info);
        int minColor = arguments.GetInt(0);
        int midColor = arguments.GetInt(1);
        int maxColor = arguments.GetInt(2);
        int minType = arguments.GetInt(3, libxl::CFVO_FORMULA);
        CSNanUtf8Value(minValue, arguments.GetString(4, ""));
        int midType = arguments.GetInt(5, libxl::CFVO_FORMULA);
        CSNanUtf8Value(midValue, arguments.GetString(6, ""));
        int maxType = arguments.GetInt(7, libxl::CFVO_FORMULA);
        CSNanUtf8Value(maxValue, arguments.GetString(8, ""));
        bool stopIfTrue = arguments.GetBoolean(9, false);
        ASSERT_ARGUMENTS(arguments);

        ConditionalFormatting* that = FromJS(info.This());
        ASSERT_THIS(that);

        that->GetWrapped()->add3ColorScaleFormulaRule(
            static_cast<libxl::Color>(minColor), static_cast<libxl::Color>(midColor),
            static_cast<libxl::Color>(maxColor), static_cast<libxl::CFVOType>(minType),
            info[4]->IsUndefined() ? nullptr : *minValue, static_cast<libxl::CFVOType>(midType),
            info[6]->IsUndefined() ? nullptr : *midValue, static_cast<libxl::CFVOType>(maxType),
            info[8]->IsUndefined() ? nullptr : *maxValue, stopIfTrue);
        info.GetReturnValue().Set(info.This());
    }

    void ConditionalFormatting::Initialize(Local<Object> exports) {
        Nan::HandleScope scope;

        Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(util::StubConstructor);
        t->SetClassName(Nan::New<String>("ConditionalFormatting").ToLocalChecked());
        t->InstanceTemplate()->SetInternalFieldCount(1);

        BookHolder::Initialize<ConditionalFormatting>(t);

        Nan::SetPrototypeMethod(t, "addRange", AddRange);
        Nan::SetPrototypeMethod(t, "addRule", AddRule);
        Nan::SetPrototypeMethod(t, "addTopRule", AddTopRule);
        Nan::SetPrototypeMethod(t, "addOpNumRule", AddOpNumRule);
        Nan::SetPrototypeMethod(t, "addOpStrRule", AddOpStrRule);
        Nan::SetPrototypeMethod(t, "addAboveAverageRule", AddAboveAverageRule);
        Nan::SetPrototypeMethod(t, "addTimePeriodRule", AddTimePeriodRule);
        Nan::SetPrototypeMethod(t, "add2ColorScaleRule", Add2ColorScaleRule);
        Nan::SetPrototypeMethod(t, "add2ColorScaleFormulaRule", Add2ColorScaleFormulaRule);
        Nan::SetPrototypeMethod(t, "add3ColorScaleRule", Add3ColorScaleRule);
        Nan::SetPrototypeMethod(t, "add3ColorScaleFormulaRule", Add3ColorScaleFormulaRule);

        t->ReadOnlyPrototype();
        constructor.Reset(Nan::GetFunction(t).ToLocalChecked());
        Nan::Set(exports, Nan::New<String>("ConditionalFormatting").ToLocalChecked(),
                 Nan::New(constructor));
    }
}  // namespace node_libxl
