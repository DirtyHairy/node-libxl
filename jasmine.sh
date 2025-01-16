#!/bin/bash

SYSTEM=$(uname)

if [ "$SYSTEM" = Darwin ]; then
    DYLD_LIBRARY_PATH="$(pwd)/deps/libxl/lib" node ./node_modules/.bin/jasmine --config=jasmine.json
else
    echo unsupported system
fi