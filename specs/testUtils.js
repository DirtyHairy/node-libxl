var path = require('path'),
    fs = require('fs');

var outputDir = path.join(__dirname, 'output'),
    writeTestFile = path.join(outputDir, 'writetest.xls');

module.exports = {
    initFilesystem: function() {
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        if (fs.existsSync(writeTestFile)) {
            fs.unlinkSync(writeTestFile);
        }
    },

    getWriteTestFile: function() {
        return writeTestFile;
    },

    shouldThrow: function(fun, scope) {
        var args = Array.prototype.slice.call(arguments, 2);

        expect(function() {fun.apply(scope, args);}).toThrow();
    }
};
