var os = require('os'),
    path = require('path');

var bindings = null;

var paths = process.env['NODE_LIBXL_PATH'] ? [process.env['NODE_LIBXL_PATH']] : [];

paths.push(
    path.join(__dirname, os.platform() + '-' + os.arch()),
    path.join(__dirname, os.platform()),
    path.join(__dirname, '..', 'build', 'Debug'),
    path.join(__dirname, '..', 'build', 'Release')
);

for (var i = 0; i < paths.length; i++) {
    try {
        bindings = require(path.join(paths[i], 'libxl'));
        break;
    } catch (e) {}
}

if (!bindings) {
    throw new Error('unable to load libxl.node');
}

module.exports = bindings;
