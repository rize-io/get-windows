const path = require('path');
const bindingPath = path.join(__dirname, 'windows-binding.js');
const addon = __non_webpack_require__(bindingPath);

module.exports = async () => addon.getActiveWindow();

module.exports.getOpenWindows = async () => addon.getOpenWindows();

module.exports.sync = addon.getActiveWindow;

module.exports.getOpenWindowsSync = addon.getOpenWindows;
