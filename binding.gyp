{
  'targets': [
    {
      'target_name': 'libxl_bindings',
      'sources': [
        'src/bindings.cc',
        'src/book.cc',
        'src/argument_helper.cc',
        'src/util.cc'
      ],
      'include_dirs': [
        'deps/libxl/include_cpp'
      ],
      'conditions': [
        ['OS=="linux"', {
          'target_name': 'liblibxl',
          'conditions': [
            ['target_arch=="ia32"', {
              'link_settings': {
                'ldflags': ['-L../deps/libxl/lib']
              }
            }],
            ['target_arch=="x64"', {
              'link_settings': {
                'ldflags': ['-L../deps/libxl/lib64']
              }
            }]
          ],
          'libraries': [
            '-lxl'
          ]
        }],
        ['OS=="win"', {
          'libraries': [
            '../deps/libxl/lib/libxl.lib'
          ]
        }]
      ]
    }
  ]
}
