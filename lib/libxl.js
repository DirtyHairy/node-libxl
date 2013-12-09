var os = require('os'),
    path = require('path');

var bindings = null;

var paths = [
    path.join(__dirname, os.platform() + '-' + os.arch(), 'libxl'),
    path.join(__dirname, os.platform(), 'libxl'),
    path.join(__dirname, '..', 'build', 'Debug', 'libxl'),
    path.join(__dirname, '..', 'build', 'Release', 'libxl')
];

for (var i = 0; i < paths.length; i++) {
    try {
        bindings = require(paths[i]);
        break;
    } catch (e) {}
}

if (!bindings) {
    throw new Error('unable to load libxl.node');
}

module.exports = bindings;
