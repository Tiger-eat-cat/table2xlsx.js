"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.imgProcessor = exports.inputProcessor = exports.hyperlinkProcessor = exports.columnProcessor = exports.fontProcessor = void 0;
var tools_1 = require("./tools");
var config_1 = require("./config");
var types_1 = require("./types");
var fontProcessor = function (cell, sheetCell, style) {
    var fontSize = style.fontSize;
    var textAlign = style.textAlign;
    var color = tools_1.rgbToArgb(style.color);
    var fontWeight = style.fontWeight;
    var BOLD = 700;
    sheetCell.font = {
        size: parseInt(fontSize),
        color: { argb: color },
        italic: style.fontStyle === 'italic',
        bold: fontWeight === 'bold' || parseInt(fontWeight) >= BOLD
    };
    if (config_1.TEXT_ALIGN.some(function (item) { return item === textAlign; })) {
        sheetCell.alignment = {
            // @ts-ignore
            horizontal: textAlign,
        };
    }
};
exports.fontProcessor = fontProcessor;
var columnProcessor = function (worksheet, from, to, cellStyle) {
    var CONVERT_RATIO = 0.35;
    if (from === to && cellStyle.width !== 'auto') {
        worksheet.getColumn(from).width = parseFloat(cellStyle.width) * CONVERT_RATIO;
    }
};
exports.columnProcessor = columnProcessor;
var hyperlinkProcessor = function (cell, sheetCell) {
    var _a;
    var children = cell.children;
    var tagName = (_a = children[0]) === null || _a === void 0 ? void 0 : _a.tagName.toUpperCase();
    if (tagName === types_1.TagName.hyperlink) {
        var hyperlink = children[0];
        sheetCell.value = { text: cell.innerText, hyperlink: hyperlink.href };
    }
};
exports.hyperlinkProcessor = hyperlinkProcessor;
var inputProcessor = function (cell, sheetCell) {
    var _a;
    var children = cell.children;
    var tagName = (_a = children[0]) === null || _a === void 0 ? void 0 : _a.tagName.toUpperCase();
    if (tagName === types_1.TagName.input) {
        var input = children[0];
        sheetCell.value = input.value;
    }
};
exports.inputProcessor = inputProcessor;
var imgProcessor = function (cell, sheetCell) {
    var _a;
    var children = cell.children;
    var tagName = (_a = children[0]) === null || _a === void 0 ? void 0 : _a.tagName.toUpperCase();
    if (tagName === types_1.TagName.img) {
        var img = children[0];
        sheetCell.value = { text: img.src, hyperlink: img.src };
    }
};
exports.imgProcessor = imgProcessor;
