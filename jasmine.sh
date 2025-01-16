#!/bin/bash

SYSTEM=$(uname)
MACHINE=$(uname -m)
JASMINE="node ./node_modules/.bin/jasmine --config=jasmine.json"

if [ "$SYSTEM" = Darwin ]; then
    DYLD_LIBRARY_PATH="$(pwd)/deps/libxl/lib" $JASMINE
elif [ "$SYSTEM" = Linux ] && [ "$MACHINE" = aarch64 ]; then
    LD_LIBRARY_PATH="$(pwd)/deps/libxl/lib-aarch64" $JASMINE
else
    echo unsupported system
fi
