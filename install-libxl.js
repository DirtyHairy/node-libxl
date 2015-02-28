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

VERSION = '3.6.1';

var fs = require('fs'),
    os = require('os'),
    path = require('path'),
    util = require('util'),
    request = require('request'),
    tar = require('tar'),
    zlib = require('zlib');

var dependencyDir = 'deps',
    libxlDir = path.join(dependencyDir, 'libxl');

function getPlatform() {
  if (os.platform().match(/^win/)) {
    return 'win';
  }else if (os.platform().match(/^linux/)) {
    return 'lin';
  }else if (os.platform().match(/^darwin/)) {
    return 'mac';
  }

  throw new Error('Unsupported platform');
}

function download() {
  var file = util.format('libxl-%s-%s', getPlatform(), VERSION);
  var url = util.format('http://www.libxl.com/download/%s.tar.gz', file);
  // the env has priority
  url = process.env.LIBXL_DOWNLOAD_URL || url;
  console.log('Downloading ' + url);
  request.get(url)
    .pipe(zlib.createGunzip())
    .pipe(tar.Extract({
      path: libxlDir,
      strip: 1
    }))
    .on('error', function(err) {
      throw err;
    })
    .on('end', function() {
      console.log('All done!');
    });
}

if (fs.existsSync(libxlDir)) {
  console.log('Libxl already downloaded, nothing to do');
  process.exit(0);
}

if (!fs.existsSync(dependencyDir)) {
  fs.mkdirSync(dependencyDir);
}

download();
