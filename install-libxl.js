/**
 * The MIT License (MIT)
 * 
 * Copyright (c) 2013 Christian Speckner <cnspeckn@googlemail.com>,
 *                    Torben Fitschen <teddyttn@gmail.com>
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */


var fs = require('fs'),
    http = require('http'),
    os = require('os'),
    path = require('path'),
    spawn = require('child_process').spawn;

var isWin = !!os.platform().match(/^win/),
    isMac = !!os.platform().match(/^darwin/),
    dependencyDir = 'deps',
    libxlDir = path.join(dependencyDir, 'libxl'),
    archiveUrl = 'http://www.libxl.com/download/libxl' + (isMac ? '-mac' : '') + (isWin ? '.zip' : '.tar.gz'),
    archiveFile = path.join(dependencyDir, path.basename(archiveUrl));

var download = function(url, file, callback) {
  var writer = fs.createWriteStream(file);
  http.get(url, function(response) {
    response.pipe(writer);
    response.on('end', function() {
      callback();
    });
  });
};
var execute = function(cmd, args, callback) {
  spawn(cmd, args).on('close', function(code) {
    if (0 === code) {
      callback();
    } else {
      process.exit(1);
    }
  });
};
var extractor = function(file, target, callback) {
  if (isWin) {
    execute(path.join('tools', '7zip', '7za.exe'), ['x', archiveFile, '-o' + dependencyDir], callback);
  } else {
    execute('tar', ['-C', dependencyDir, '-zxf', archiveFile], callback);
  }
};
var finder = function(dir, pattern) {
  var files = fs.readdirSync(dir),
      i,
      file;
  for (i = 0; i < files.length; i++) {
    file = files[i];
    if (file.match(pattern)) {
      return path.join(dir, file);
    }
  }
  return null;
};

if (fs.existsSync(libxlDir)) {
  process.exit(0);
}

if (!fs.existsSync(dependencyDir)) {
  fs.mkdirSync(dependencyDir);
}

download(archiveUrl, archiveFile, function() {
  extractor(archiveFile, dependencyDir, function() {
    fs.unlinkSync(archiveFile);
    fs.renameSync(finder(dependencyDir, /^libxl/), libxlDir);
  });
});
