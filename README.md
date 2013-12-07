# What it is

Node.js bindings for [libxl](http://www.libxl.com/).

# How to use it

Pull the library into your project with `npm instll libxl`
and require the module via

    var xl = require('libxl');

A new excel document is then created via

    var xlsBook = new xl.Book(xl.BOOK_TYPE_XLS);

or

    var xlsxBook = new xl.Book(xl.BOOK_TYPE_XLSX);

(for xlsx documents). The API closely follows the
[libxl documentation](http://www.libxl.com/documentation.html).
For example, adding a new sheet and writing two cells works as

    var sheet = xlsBook.addShet('Sheet 1');
    sheet.writeStr(1, 0, 'A string');
    sheet.writeNum(1, 1, 42);

Functions whose C++ counterpart returns void or an error status
have been implemented to return the respective instance, so it
is possible to chain calls

    sheet
        .writeStr(1, 0, 'A string');
        .writeNum(1, 1, 42);

Errors are handled by throwing exceptions.

API coverage is a work in progress; see demo.js,
jasmine specs and the class initializers
(Book::Initialize, Sheet::Initialize, Format::Initialize)
to find out what is currently supported :).
