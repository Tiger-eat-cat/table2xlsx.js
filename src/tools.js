"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.rgbToArgb = void 0;
var rgbToArgb = function (rgb) {
    var numbers = rgb.split('(')[1].split(')')[0].split(',')
        .map(function (number, index) {
        return index === 3 ? "" + parseInt(number) * 255 : number;
    });
    if (numbers.length === 3) {
        numbers.unshift('255');
    }
    var argb = numbers.map(function (number) {
        var hex = parseInt(number).toString(16);
        return hex.length === 1 ? "0" + hex : hex;
    });
    return argb.join('').toUpperCase();
};
exports.rgbToArgb = rgbToArgb;
