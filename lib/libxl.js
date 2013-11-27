var binding;

try {
	binding = require(__dirname + '/../build/Debug/libxl');
} catch(e) {
	binding = require(__dirname + '/../build/Release/libxl');
}

module.exports = binding;
