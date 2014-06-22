{
  'targets': [
    {
      'target_name': 'libxl',
      'sources': [
        'src/bindings.cc',
        'src/book.cc',
        'src/argument_helper.cc',
        'src/util.cc',
        'src/sheet.cc',
        'src/format.cc',
        'src/font.cc',
        'src/book_wrapper.cc',
        'src/string_copy.cc',
        'src/worker.cc',
        'src/buffer_copy.cc'
      ],
      'include_dirs': [
        'deps/libxl/include_cpp',
        "<!(node -e \"require('nan')\")"
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
        }],
        ['OS=="mac"', {
          'target_name': 'liblibxl',
          'link_settings': {
            'libraries': [
              '-I../deps/libxl/include_cpp',
              '-L../deps/libxl/lib',
              '-lxl'
            ]
          }
        }]
      ]
    }
  ]
}
