var path = require('path'),
    fs = require('fs');

var outputDir = path.join(__dirname, 'output'),
    writeTestFile = path.join(outputDir, 'writetest.xls'),
    tempFile = path.join(outputDir, 'tempfile'),
    filesDir = path.join(__dirname, 'files'),
    testPicture = path.join(filesDir, 'dummy.jpg');

module.exports = {
    initFilesystem: function () {
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        if (fs.existsSync(writeTestFile)) {
            fs.unlinkSync(writeTestFile);
        }
    },

    getWriteTestFile: function () {
        return writeTestFile;
    },

    getTempFile: function () {
        return tempFile;
    },

    shouldThrow: function (fun, scope) {
        var args = Array.prototype.slice.call(arguments, 2);

        expect(function () {
            fun.apply(scope, args);
        }).toThrow();
    },

    getTestPicturePath: function () {
        return testPicture;
    },

    compareBuffers: function (buf1, buf2) {
        if (buf1.length !== buf2.length) return false;

        for (var i = 0; i < buf1.length; i++) {
            if (buf1[i] !== buf2[i]) return false;
        }

        return true;
    },

    testPictureWidth: 640,
    testPictureHeight: 480,
};
