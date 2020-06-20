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
    childProcess = require('child_process'),
    os = require('os'),
    path = require('path'),
    tmp = require('tmp'),
    util = require('util'),
    md5 = require('md5'),
    zlib = require('zlib'),
    tar = require('tar'),
    plist = require('plist'),
    https = require('https'),
    AdmZip = require('adm-zip');

var isWin = !!os.platform().match(/^win/),
    isMac = !!os.platform().match(/^darwin/),
    dependencyDir = 'deps',
    libxlDir = path.join(dependencyDir, 'libxl'),
    archiveEnv = 'NODE_LIBXL_SDK_ARCHIVE';

function download(callback) {
    function getPlaform() {
        if (isWin) {
            return 'win';
        }

        if (isMac) {
            return 'mac';
        }

        return 'lin';
    }

    function getArchiveSuffix() {
        switch (getPlaform()) {
            case 'win':
                return 'zip';

            case 'mac':
                return 'dmg';

            default:
                return 'tar.gz';
        }
    }

    function getArchiveName() {
        return util.format('libxl-%s-latest.%s', getPlaform(), getArchiveSuffix());
    }

    function downloadToFile(filename, callback) {
        var writer;

        function dieOnError(e) {
            console.log(util.format('unable to download the libxl SDK: %s', e.message));
            console.log(util.format(
                '\nplease download libxl manually from https://www.libxl.com and point the environment variable %s to the downloaded file',
                archiveEnv
            ));

            process.exit(1);
        }

        function onOpen() {
            var url = util.format('https://www.libxl.com/download/%s', getArchiveName());

            console.log('Downloading ' + url);

            https.get(url, function(response) {
                if (response.statusCode !== 200) {
                    dieOnError(new Error(util.format('request failed: %s %s', response.statusCode, response.statusMessage)));
                }

                response.on('error', dieOnError);

                writer.on('finish', function() {
                    callback(filename);
                });

                response.pipe(writer);
            }).on('error', dieOnError);
        }

        writer = fs.createWriteStream(filename);
        writer.on('error', dieOnError);
        writer.on('open', onOpen);
    }

    function calculateMd5(filename, callback) {
        fs.readFile(filename, function(err, buffer) {
            if (err) {
                throw err;
            }

            callback(md5(buffer));
        });
    }

    function onDownloadComplete(archive) {
        calculateMd5(archive, function(md5) {
            console.log(util.format('successfully downloaded %s, MD5: %s', archive, md5));

            callback(archive);
        });
    }

    tmp.tmpName({
        postfix: path.basename(getArchiveName()),
        tries: 10
    }, function(err, outfile) {
        downloadToFile(outfile, onDownloadComplete);
    });
};

function downloadIfNecessary(callback) {
    var suppliedArchive = process.env[archiveEnv];

    if (suppliedArchive) {
        console.log(util.format('Automatic download overriden by %s, using archive "%s"...',
            archiveEnv, suppliedArchive
        ));
        callback(suppliedArchive);
    } else {
        download(callback);
    }
};

function findSubdir(dir, pattern) {
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

function extractTgz(archive, destination, callback) {
    var fileStream = fs.createReadStream(archive),
        decompressedStream = fileStream.pipe(zlib.createGunzip()),
        untarStream = tar.Extract({path: destination});

    untarStream.on('end', function() {
        callback();
    });

    [fileStream, decompressedStream, untarStream].forEach(function(stream) {
        stream.on('error', function(e) {
         callback(e);
        });
    });

    decompressedStream.pipe(untarStream);
};

function extractZip(archive, destination, callback) {
    var zip;

    try {
        zip = new AdmZip(archive);
        zip.extractAllTo(destination);

        callback();
    } catch (e) {
        callback(e);
    }
};

function extractDmg(image, destination, callback) {
    childProcess.execFile('hdiutil', ['mount', '-plist', image], {encoding: 'utf8'}, function(err, stdout) {
        var result, candidates, i, mountpoint, sdkdir;

        function unmount(cb) {
            var tries = 0;

            function work() {
                childProcess.execFile('hdiutil', ['unmount', mountpoint], function(err) {
                    if (err && tries++ < 5) {
                        setTimeout(work, 1000);
                    } else {
                        cb(err);
                    }
                });
            }

            if (mountpoint) {
                work();
            } else {
                cb();
            }
        }

        if (err) {
            callback(new Error(util.format("failed to mount %s: %s", image, err.message)));
            return;
        }

        try {
            result = plist.parse(stdout);

            candidates = result['system-entities'];
            if (!candidates) throw new Error();

            for (i = 0; i < candidates.length; i++) {
                if (mountpoint = candidates[i]['mount-point']) break;
            }

            if (!mountpoint) throw Error();
        } catch (e) {
            callback(new Error('failed to identify dmg mountpoint'));

            return;
        }

        sdkdir = findSubdir(mountpoint, /^libxl/);
        if (!sdkdir) {
            unmount(function() {
                callback(new Error('SDK not found in dmg'));
            });

            return;
        }

        childProcess.execFile('cp', ['-r', findSubdir(mountpoint, /^libxl/), destination], function(err) {
            unmount(function() {
                callback(err);
            });
        });
    })
}

function extractArchive(file, target, callback) {
    console.log('Extracting ' + file + ' ...');

    if (file.match(/\.zip$/)) {
        extractZip(file, target, callback);
    } else if (file.match(/\.tar\.gz/)) {
        extractTgz(file, target, callback);
    } else if (file.match(/\.dmg$/)) {
        extractDmg(file, target, callback);
    } else {
        callback(new Error('unnown archive format'));
    }
};

if (fs.existsSync(libxlDir)) {
    console.log('Libxl already downloaded, nothing to do');
    process.exit(0);
}

if (!fs.existsSync(dependencyDir)) {
    fs.mkdirSync(dependencyDir);
}

downloadIfNecessary(function(archive) {
    extractArchive(archive, dependencyDir, function(e) {
        if (e) {
            console.error(e.message || 'Extraction failed');
            process.exit(1);
        }

        if (!process.env[archiveEnv]) fs.unlinkSync(archive);

        var extractedDir = findSubdir(dependencyDir, /^libxl/);
        console.log('Renaming ' + extractedDir + ' to ' + libxlDir + ' ...');

        fs.renameSync(extractedDir, libxlDir);

        console.log('All done!');
    });
});
