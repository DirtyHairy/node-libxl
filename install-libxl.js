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
    Ftp = require('ftp'),
    os = require('os'),
    path = require('path'),
    spawn = require('child_process').spawn;

var isWin = !!os.platform().match(/^win/),
    isMac = !!os.platform().match(/^darwin/),
    dependencyDir = 'deps',
    libxlDir = path.join(dependencyDir, 'libxl'),
    ftpHost = 'xlware.com';

var download = function(callback) {
    var ftpClient = new Ftp();

    function decodeDirectoryEntry(entry) {
        var match = entry.name.match(/^libxl-(\w+)-([\d\.]+)\.([a-zA-z\.]+)$/);
        if (match) {
            return {
                file: entry.name,
                system: match[1],
                version: match[2],
                suffix: match[3]
            };
        }
    }

    function decodeVersion(file) {
        return file.version.split('.');
    }

    function compareFiles(file1, file2) {
        function cmp(a, b) {
            if (a > b) return -1;
            if (a < b) return  1;
            return 0;
        }

        var v1 = decodeVersion(file1), v2 = decodeVersion(file2),
            nibbles = Math.min(v1.length, v2.length),
            partialResult;

        for (var i = 0; i < nibbles; i++) {
            partialResult = cmp(v1[i], v2[i]);

            if (partialResult) return partialResult;
        }

        return v1.length > nibbles ? -1 : 1;
    }

    function validArchive(file) {
        if (isWin) {
            return file.system === "win" && file.suffix === "zip";
        } else if (isMac) {
            return file.system === "mac" && file.suffix === "tar.gz";
        }
        return file.system === "lin" && file.suffix === "tar.gz";
    }

    function onError(error) {
        console.log('Download from FTP failed');
        throw(error);
    }

    function onReady() {
        ftpClient.list(function(error, list) {
            if (error) onError(error);

            console.log('Connected, receiving directory list...');
            processDirectoryList(list);
        });
    }

    function processDirectoryList(list) {
        try {
            if (!list) throw new Error('FTP list failed');

            var candidates = list
                .map(decodeDirectoryEntry)
                .filter(validArchive)
                .sort(compareFiles);

            if (!candidates.length) throw new Error('Failed to identify a suitable download');

            download(candidates[0].file);
        } catch (error) {
            console.log(error.message);
        }
    }

    function download(name) {
        console.log('Downloading ' + name + '...');

        ftpClient.get(name, function(error, stream) {
            var outfile = path.join(dependencyDir, name),
                writer = fs.createWriteStream(outfile);

            writer.on('error', onError);
            stream.on('error', onError);
            writer.on('close', function() {
                ftpClient.end();

                console.log('Download complete!');

                callback(outfile);
            });

            stream.pipe(writer);
        });
    }

    ftpClient.on('error', onError);
    ftpClient.on('ready', onReady);

    console.log('Connecting to ftp://' + ftpHost + '...');

    ftpClient.connect({
        host: ftpHost
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
    console.log('Extracting ' + file + ' ...');

    if (isWin) {
        execute(path.join('tools', '7zip', '7za.exe'), ['x', file, '-o' + dependencyDir], callback);
    } else {
        execute('tar', ['-C', dependencyDir, '-zxf', file], callback);
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

download(function(archive) {
    extractor(archive, dependencyDir, function() {
        fs.unlinkSync(archive);

        var extractedDir = finder(dependencyDir, /^libxl/);
        console.log('Renaming ' + extractedDir + ' to ' + libxlDir + ' ...');

        fs.renameSync(extractedDir, libxlDir);

        console.log('All done!');
    });
});
