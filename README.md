[![Build Status](https://travis-ci.org/DirtyHairy/node-libxl.svg?branch=release-0.2)](https://travis-ci.org/DirtyHairy/node-libxl)
[![npm version](https://badge.fury.io/js/libxl.svg)](https://badge.fury.io/js/libxl)
# What it is

Node.js bindings for [libxl](http://www.libxl.com/). Node versions starting from 10.0.0
are supported (see 'Platform and Node.js support' below for details).

# Compilation and Installation

Pull the library into your project with `npm install libxl`
and require the module via

    var xl = require('libxl');

## LibXL installation

As this packages contains only bindings for the libxl library, the library
itself is required for building and running the bindings.

### Compilation Phase

Before the bindings are compiled, the `install-libxl.js` script pulls the latest
version of the library from the XLware FTP server and unpacks it in `deps/libxl`.
Therefore, **no separate installation of libxl is necessary for building the
bindings**.

If you want to compile the bindings against a particular version of the library
or if, for some reason, the automatic download fails, you can point the install
script to a locally downloaded archive of the SDK by setting the
`NODE_LIBXL_SDK_ARCHIVE` environment variable.

### Runtime

In order to load and use the bindings, the libxl library must be available in
your dynamic library search path. This is achieved by either

**Copying the library into your system library search path**, e.g. `/usr/lib` on
  Linux.

**Copying the library into the working directory** where you run the script which
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

or

    xlsBook.write('file.xls', callback);

and read back via

    xlsBook.loadSync('file.xls');

or

    xlsBook.load('file.xls', callback);

where `callback` will be called after the operation has completed, receiving an
optional error object as argument if anything goes wrong.

**IMPORTANT:** See below for additional notes on the async implementation
of libxl calls.

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

Functions that return multiple values by reference in C++ (like
Book::dateUnpack) usually return a object with the return values as properties.

If in doubt on how to call a method, please check the specs in the `spec`
directory for invocation examples. Every methods is covered with at least one tests.

**IMPORTANT:** The Javascript API enforces the types defined in its C++
counterpart for all function arguments; there is no implicit type casting. For
example, passing a number to Sheet::writeStr will throw a TypeError instead of
silently converting the string to a number.

## Asynchroneous variants of libxl calls

The async variants of libxl calls implement the standard Node.js API for
async functions: a callback is passed as last argument which is called once the
operation has finished. The first argument of the callback is an error object,
which is `undefined` if the operation completed without errors. Any results are
passed as additional arguments to the callback.

**IMPORTANT:** While an async operation is pending, other operations (sync or
async) on the same book object (and its descendants like sheets, formats and
fonts) are not allowed and will throw an exception. However, multiple
simultaneous operations on different books are allowed.

The following async functions are available:

* `book.write` / `book.save`, `book.load` are implemented asynchroneously. If
  you need synchroneous behavior you can use `book.loadSync` etc.
* `book.writeRaw` / `book.saveRaw`, `book.loadRaw` are implemented
  asynchroneously. `book.saveRaw` and its alias return the book data as second
  argument to the supplied callback. Use `book.loadRawSync` & friends for
  synchroneous behavior.
* `book.addPicture` has a async version `book.addPictureAsync`. The index of the
  new picture is passed as the second argument to the callback.
* `book.getPicture` has a async version `book.getPictureAsync`. Picture type and
  data are passed to the callback as second and third arguments.
* `sheet.insertRow` and `sheet.insertCol` are very slow and thus are also
  available as async implementations `sheet.insertRowAsync` and
  `sheet.insertColAsync`.

## Other differences

* Book object creation: Books are **not** created via `xlCreateBook` and
  `xlCreateXMLBook`. Instead, object instances are directly constructed from the
  `xl.Book` constructor via either `new xl.Book(xl.BOOK_TYPE_XLS)` or `new xl.Book(xl.BOOK_TYPE_XLSX)`
* Accessing the parent book: all objects hold a reference to
  their parent book that can be accessed via the `book` property

## Enum constants

All C enum constants provided by the library are available as constants on the
library object, e.g.  `xl.NUMFORMAT_DATE` or `xl.PICTURETYPE_PNG`.

## Unlocking the API

If you have purchased a licence key from XLware, you can call book.setKey in
order to unlock the library. As an alternative, you can build the key into the
bindings by modifying `api_key.h` and rebuilding the library via `node-gyp
rebuild` (you'll have to install node-gyp for this) or `npm install` in the
package directory.

## Coverage

The bindings cover the current (version 4.5.0) libxl API completely.

# Platform and Node.js support

## Platforms

The package supports Linux, Windows and Mac.

## Node.js

The current branch (0.6.x) supports all current versions of Node.js starting with 18.0.0.
If you need support for older Node versions you'll have to use an older version of the
bindings.

Be aware, though, that the build time dependencies have moved on and are starting to use
ES6 features that break support for Node < 6, even if those older versions of node-libxl
still support it. If you still need to use such an old version of Node, you'll have to
use lockfiles to pin those transitive dependencies.

## Restrictions

The async hooks API introduced with Node.js 9 is currently unsupported.

# Tests

The bindings are fully covered with jasmine tests. If you have jasmine-node
installed (via NPM), you can run the suite via

    jasmine-node specs/

# Reporting bugs

Please report any bugs or feature requests on the github issue tracker.

# Roadmap

As the API is completely covered, I consider the bindings complete. New releases
will only cover new libxl methods and fix bugs. If you identify parts of
libxl that are particularily slow, asynchroneous version of those could be added
as well. Note that only the latest version branch (0.6.x) is maintained and
supported.

# Credits

Developed with kind support from XLware.

* Torben Fitschen wrote the install script which pulls the
  necessary libxl SDK before building.
* Martin Schröder for adding Mac support.
* Parts if this package were developed during slacktime provided by the awesome folks at
  [Mayflower GmbH](http://www.mayflower.de)
* Alexander Makarenko wrote
  [node-excel-libxl](https://github.com/7eggs/node-excel-libxl)
  Though node-libxl is rewritten from scratch, this
  package served as the starting point.
