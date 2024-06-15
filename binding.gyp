{
    "targets": [
        {
            "target_name": "xlsx",
            "sources": ["xlsx.cpp", "xlsx.h"],
            "include_dirs" : ["./include/"],
            "libraries": ["<(module_root_dir)/include/build/Release/OpenXLSX"],
            "conditions": [
                ['OS=="win"', {
                  'configurations': {
                    'Release': {
                      'defines': [ 'NDEBUG' ],
                      'msvs_settings': {
                        'VCCLCompilerTool': { 
                            'Optimization': 'MaxSpeed',
                            "RuntimeLibrary": 2
                        },
                        'VCLinkerTool': { 'EnableCOMDATFolding': '1', 'OptimizeReferences': '2' }
                      }
                    }
                  }
                }],
                ['OS!="win"', {
                  'configurations': {
                    'Release': {
                      'cflags!': [ '-g' ],
                      'cflags_cc!': [ '-g' ],
                      'ldflags!': [ '-g' ]
                    }
                  }
                }]
            ]
        }
    ]
}