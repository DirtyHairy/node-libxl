# This file is NOT licensed under the GPLv3, which is the license for the rest
# of YouCompleteMe.
#
# Here's the license text for this file:
#
# This is free and unencumbered software released into the public domain.
#
# Anyone is free to copy, modify, publish, use, compile, sell, or
# distribute this software, either in source code form or as a compiled
# binary, for any purpose, commercial or non-commercial, and by any
# means.
#
# In jurisdictions that recognize copyright laws, the author or authors
# of this software dedicate any and all copyright interest in the
# software to the public domain. We make this dedication for the benefit
# of the public at large and to the detriment of our heirs and
# successors. We intend this dedication to be an overt act of
# relinquishment in perpetuity of all present and future rights to this
# software under copyright law.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
# OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
# ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.
#
# For more information, please refer to <http://unlicense.org/>

import os
import ycm_core

# These are the compilation flags that will be used in case there's no
# compilation database set (by default, one is not set).
# CHANGE THIS LIST OF FLAGS. YES, THIS IS THE DROID YOU HAVE BEEN LOOKING FOR.
flags = [
'-Wall',
'-Wc++98-compat',
'-Wno-long-long',
'-Wno-variadic-macros',
# For a C project, you would set this to 'c' instead of 'c++'.
'-x',
'c++',
'-isystem',
'../BoostParts',
'-isystem',
# This path will only work on OS X, but extra paths that don't exist are not
# harmful
'/System/Library/Frameworks/Python.framework/Headers',
'-isystem',
'../llvm/include',
'-isystem',
'../llvm/tools/clang/include',
'-I',
'.',
'-I',
'./deps/libxl/include_cpp',
'-I',
'/home/pestix/.node-gyp/0.10.20/src',
'-I',
'/home/pestix/.node-gyp/0.10.20/deps/uv/include',
'-I',
'/home/pestix/.node-gyp/0.10.20/deps/v8/include'
]

SOURCE_EXTENSIONS = [ '.cpp', '.cxx', '.cc', '.c', '.m', '.mm' ]

def DirectoryOfThisScript():
  return os.path.dirname( os.path.abspath( __file__ ) )


def MakeRelativePathsInFlagsAbsolute( flags, working_directory ):
  if not working_directory:
    return list( flags )
  new_flags = []
  make_next_absolute = False
  path_flags = [ '-isystem', '-I', '-iquote', '--sysroot=' ]
  for flag in flags:
    new_flag = flag

    if make_next_absolute:
      make_next_absolute = False
      if not flag.startswith( '/' ):
        new_flag = os.path.join( working_directory, flag )

    for path_flag in path_flags:
      if flag == path_flag:
        make_next_absolute = True
        break

      if flag.startswith( path_flag ):
        path = flag[ len( path_flag ): ]
        new_flag = path_flag + os.path.join( working_directory, path )
        break

    if new_flag:
      new_flags.append( new_flag )
  return new_flags



def FindNodeGypPath():

  basepath = os.path.expanduser(os.path.join("~", ".node-gyp"));

  def GetLarger(v1, v2):
    v1_comp = v1.split(".")
    v2_comp = v2.split(".")

    for i in range(0, min(len(v1_comp), len(v2_comp))):
      try:
        comp1 = int(v1_comp[i])
      except:
        comp1 = -1

      try:  
        comp2 = int(v2_comp[i])
      except:
        comp1 = -1

      if comp1 < comp2:
        return v2
      if comp1 > comp2:
        return v1

    return v1

  def FindBest(candidates):
    if (len(candidates) == 0):
      return None
    best = candidates[0]

    for candidate in candidates:
      best = GetLarger(best, candidate)

    return best

  return FindBest(os.listdir(basepath))

  
def FlagsForFile( filename, **kwargs ):
  relative_to = DirectoryOfThisScript()

  node_gyp_path = FindNodeGypPath()

  final_flags = flags

  if node_gyp_path:
    final_flags += ["-I", os.path.join(node_gyp_path, "src")]
    final_flags += ["-I", os.path.join(node_gyp_path, "deps", "uv", "include")]
    final_flags += ["-I", os.path.join(node_gyp_path, "deps", "v8", "include")]

  final_flags = MakeRelativePathsInFlagsAbsolute( flags, relative_to )

  return {
    'flags': final_flags,
    'do_cache': True
  }
