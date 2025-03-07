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

#include "auto_filter.h"
#include "book.h"
#include "common.h"
#include "conditional_format.h"
#include "conditional_formatting.h"
#include "core_properties.h"
#include "enum.h"
#include "filter_column.h"
#include "font.h"
#include "form_control.h"
#include "format.h"
#include "rich_string.h"
#include "sheet.h"

using namespace v8;
using namespace node_libxl;

void Initialize(Local<Object> exports) {
    Book::Initialize(exports);
    Sheet::Initialize(exports);
    Format::Initialize(exports);
    Font::Initialize(exports);
    CoreProperties::Initialize(exports);
    RichString::Initialize(exports);
    AutoFilter::Initialize(exports);
    FilterColumn::Initialize(exports);
    FormControl::Initialize(exports);
    ConditionalFormat::Initialize(exports);
    ConditionalFormatting::Initialize(exports);
    DefineEnums(exports);
}

NODE_MODULE(libxl, Initialize)
