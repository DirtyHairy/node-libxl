# What it is

Node.js bindings for [libxl](http://www.libxl.com/). Both Node 0.10 and Node
0.11 starting with 0.11.13 are supported.

# Compilation and Installation

Pull the library into your project with `npm install libxl`
and require the module via

    var xl = require('libxl');

## LibXL installation

As this packages contains only bindings for the libxl library, the library
itself is required for building and running the bindings.

### Compilation Phase

Before the bindings are compiled, the `install-libxl.js` script pulls the latest
version of the library from the libxl website and unpacks it in `deps/libxl`.
Therefore, **no separate installation of libxl is necessary for building the
bindings**.

However, if you want to compile and run against a particular version of libxl,
you can do so by manually unpacking the library archive into `deps/libxl` before
building the bindings. This will bypass the `install-libxl.js` script and build
the bindings against that specific version of the library.

### Runtime

In order to load and use the bindings, the libxl library must be available in
your dynamic library search path. This is achieved by either

**Copying the library into your system library search path**, e.g. `/usr/lib` on
  Linux.
  
**Copying the library into the working directory** where you run the scrip which
  uses the bindings. The name of the library file is `libxl.so` on Linux,
  `libxl.dylib` on Mac and `libxl.dll` on Windows.
  
**Properly setting the `LD_LIBRARY_PATH` (Linux) or `DYLD_LIBRARY_PATH` (Mac)**
  environment variable. For example, the following command will execute the
  `demo.js` script in the package directory without requiring libxl to be
  installed separately
  
    LD_LIBRARY_PATH="`pwd`/deps/libxl/lib:`pwd`/deps/libxl/lib64:$LD_LIBRARY_PATH" node demo.js


### Overriding the location of the compiled bindings

You can override the location where the Javascript wrapper looks for the
`libxl.node` file by setting the `NODE_LIBXL_PATH` environment variable. This
allows to distribute / deploy an application that uses the bindings to a system
which runs on a different platform / architecture without recompiling the
bindings there.

# API

## Usage

A new excel document is created via

    var xlsBook = new xl.Book(xl.BOOK_TYPE_XLS);

or

    var xlsxBook = new xl.Book(xl.BOOK_TYPE_XLSX);

(for xlsx documents). The document is written to disk via

    xlsBook.writeSync('file.xls');

and read back via

    xlsBook.loadSync('file.xls');

(async operations are not implemented atm).

The Javascript API closely follows the C++ API described in the
[libxl documentation](http://www.libxl.com/documentation.html).
For example, adding a new sheet and writing two cells works as

    var sheet = xlsBook.addSheet('Sheet 1');
    sheet.writeStr(1, 0, 'A string');
    sheet.writeNum(1, 1, 42);

Functions whose C++ counterpart returns void or an error status
have been implemented to return the respective instance, so it
is possible to chain calls

    sheet
        .writeStr(1, 0, 'A string');
        .writeNum(1, 1, 42);

Errors are handled by throwing exceptions.

Functions that return multiple values
by reference in C++ (like Book::dateUnpack) return a object with the return
values as properties.

See 'Differences...' below for a more detailed description of the methods whose
behavior differs from their C++ counterpart.

**IMPORTANT:** The Javascript API enforces the types defined in its C++
counterpart for all function arguments; there is no implicit type casting. For
example, passing a Number to Sheet::writeStr will throw a TypeError instead of
silently converting the string to a number.

## Coverage

The bindings cover the libxl API almost completely. The only functions still
missing are the methods for managing pictures embedded in spreadsheets and
`book.loadRaw` / `book.saveRaw`.

## Differences w.r.t. the C++ API

### Not implemented yet

* `book.loadRaw`
* `book.saveRaw`
* `book.pictureSize`
* `book.getPicture`
* `book.addPicture`
* `book.addPicture2`
* `sheet.pictureSize`
* `sheet.getPicture`
* `sheet.addPicture`
* `sheet.addPicture2`

### Interface differences

* `book.dateUnpack`: Returns an object with `year`, `month`,
  `date`, `hour`, `minute`, `seconds` and `mseconds` properties.
* `book.colorUnpack`: Returns an object with `red`, `green` and `blue`
  properties.
* `book.defaultFont`: Returns an object with `name` and `size` properties.
* `sheet.readStr` & friends: If `sheet.readXXX` is provided with an object as
  optional second argument, the cell format is returned in the objects `format`
  property.
* `sheet.getMerge`: Returns an object with the `rowFirst`, `rowLast`,
  `colFirst`, `colLast` properties.
* `sheet.getPrintFit`: Returns either `false` or an object with the `wPages` and
  `hPages` properties.
* `sheet.getNamedRange`: Returns an object with `rowFirst`, `rowLast`,
  `colFirst`, `colLast` and `hidden` properties.
* `sheet.namedRange`: Returns an object with `rowFirst`, `rowLast`,
  `colFirst`, `colLast`, `name`, `scopeId`, and `hidden` properties.
* `sheet.getTopLeftView`: Returns an object with `row` and `col` properties.
* `sheet.addrToRowCol`: Returns an object with `row`, `col`, `rowRelative`,
  `colRelative` properties.

### Other differences

* Book object creation: Books are **not** created via `xlCreateBook` and
  `xlCreateXMLBook`. Instead, object instances are directly constructed from the
  `xl.Book` constructor via either `new xl.Book(xl.BOOK_TYPE_XLS)` or `new xl.Book(xl.BOOK_TYPE_XLSX)`
* Accessing the parent book: sheet, format and font objects hold a reference to
  their parent book that can be accessed via the `book` property

## Unlocking the API

If you have purchased a licence key from XLware, you can call book.setKey in
order to unlock the library. As an alternative, you can build the key into the
bindings by modifying `api_key.h` and rebuilding the library via `node-gyp
rebuild` (you'll have to install node-gyp for this) or `npm install` in the
package directory.

# Platform support

The package currently supports Linux, Windows and Mac.

# Tests

The bindings are fully covered with jasmine tests. If you have jasmine-node
installed (via NPM), you can run the suite via

    jasmine-node specs/

# Reporting bugs

Please report any bugs on the github issue tracker.

# Roadmap

Version 0.1 is the first release which covers nearly the full API. For 0.2 (and
possibly 0.3), I anticipate adding asynchroneous calls for reading and writing
and implementing the remaining functions.

# Credits

* Torben Fitschen wrote the install script which pulls the
  necessary libxl SDK before building.
* Martin Schröder for adding Mac support.
* Parts if this package were developed during slacktime provided by the awesome folks at
  [Mayflower GmbH](http://www.mayflower.de)
* Alexander Makarenko wrote
  [node-excel-libxl](https://github.com/7eggs/node-excel-libxl)
  Though node-libxl is rewritten from scratch, this
  package served as the starting point.
