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
                'src/book_holder.cc',
                'src/string_copy.cc',
                'src/buffer_copy.cc',
                'src/core_properties.cc',
                'src/rich_string.cc',
                'src/auto_filter.cc',
                'src/filter_column.cc',
                'src/form_control.cc',
                'src/conditional_format.cc',
                'src/conditional_formatting.cc',
                'src/enum.cc'
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
                        }],
                        ['target_arch=="arm64"', {
                            'link_settings': {
                                'ldflags': ['-L../deps/libxl/lib-aarch64']
                            }
                        }],
                        ['target_arch=="arm"', {
                            'link_settings': {
                                'ldflags': ['-L../deps/libxl/lib-armhf']
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
                    ],
                    'msvs_settings':
                    {
                        'VCLinkerTool': {
                            'AdditionalOptions': [
                                '/FORCE:MULTIPLE'
                            ]
                        }
                    }
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
