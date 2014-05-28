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
    }
};
