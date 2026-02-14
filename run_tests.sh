#!/bin/bash

SYSTEM=$(uname)
MACHINE=$(uname -m)

NODE="node"
test -n "$VALGRIND" && NODE="valgrind node"

TEST="$NODE --test --test-reporter=spec specs/*Spec.js"

if [ "$SYSTEM" = Darwin ]; then
    DYLD_LIBRARY_PATH="$DYLD_LIBRARY_PATH:$(pwd)/deps/libxl/lib" $TEST
elif [ "$SYSTEM" = Linux ] && [ "$MACHINE" = aarch64 ]; then
    LD_LIBRARY_PATH="$LD_LIBRARY_PATH:$(pwd)/deps/libxl/lib-aarch64" $TEST
elif [ "$SYSTEM" = Linux ] && [ "$MACHINE" = armv7l ]; then
    LD_LIBRARY_PATH="$LD_LIBRARY_PATH:$(pwd)/deps/libxl/lib-armhf" $TEST
elif [ "$SYSTEM" = Linux ] && [ "$MACHINE" = i686 ]; then
    LD_LIBRARY_PATH="$LD_LIBRARY_PATH:$(pwd)/deps/libxl/lib" $TEST
elif [ "$SYSTEM" = Linux ] && [ "$MACHINE" = x86_64 ]; then
    LD_LIBRARY_PATH="$LD_LIBRARY_PATH:$(pwd)/deps/libxl/lib64" $TEST
else
    echo unsupported system
fi
