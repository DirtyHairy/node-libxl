#ifndef BINDINGS_CNAN_H
#define BINDINGS_CSNAN_H

#if (NODE_MODULE_VERSION > 0x000B)

#define CSNanThrow(Error) return (void)nan_isolate->ThrowException(Error)
#define CSNanNewExternal(Value) v8::External::New(nan_isolate, Value)

#else

#define CSNanThrow(Error) return v8::ThrowException(Error)
#define CSNanNewExternal(Value) v8::External::New(Value)

#endif

#endif //BINDINGS_CSNAN_H
