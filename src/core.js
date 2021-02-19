"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createExcel = void 0;
var exceljs_1 = __importDefault(require("exceljs"));
var processor_1 = require("./processor");
var createExcel = function (selector) {
    if (selector === void 0) { selector = 'table'; }
    var workbook = new exceljs_1.default.Workbook();
    var tableElements = typeof selector === 'string' ? document.querySelectorAll(selector) : selector;
    var tables = Array.from(tableElements);
    tables.forEach(function (table, index) {
        var worksheet = workbook.addWorksheet("Worksheet" + (index + 1));
        var rows = Array.from(table.rows);
        var sheetHeight = rows.length;
        var sheetWidth = rows[sheetHeight - 1].cells.length;
        var generateRow = function () { return new Array(sheetWidth).fill(false); };
        var mergeLog = new Array(sheetHeight).fill(null).map(function () { return generateRow(); });
        rows.forEach(function (row, rowIndex) {
            var y = rowIndex + 1; // 纵坐标
            var x = 1; // 横坐标
            var currentLineLog = mergeLog[rowIndex];
            for (var i = 0; i < sheetWidth; i++) {
                if (!currentLineLog[i]) {
                    x = i + 1;
                    break;
                }
            }
            var cells = Array.from(row.cells);
            cells.forEach(function (cell) {
                var colSpan = cell.colSpan, rowSpan = cell.rowSpan;
                var top = y; // 开始行
                var left = x; // 开始列
                var bottom = y + rowSpan - 1; // 结束行
                var right = x + colSpan - 1; // 结束列
                worksheet.mergeCells(top, left, bottom, right);
                var sheetCell = worksheet.getCell(top, left);
                sheetCell.value = cell.innerText;
                var style = getComputedStyle(cell);
                processor_1.fontProcessor(cell, sheetCell, style);
                processor_1.columnProcessor(worksheet, left, right, style);
                for (var i = top - 1; i < bottom; i++) {
                    for (var j = left - 1; j < right; j++) {
                        mergeLog[i][j] = true;
                    }
                }
                x += colSpan;
            });
        });
    });
    return {
        export: function (filename) {
            if (filename === void 0) { filename = 'workbook'; }
            return __awaiter(void 0, void 0, void 0, function () {
                var buffer, a, fileUrl;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0: return [4 /*yield*/, workbook.xlsx.writeBuffer()];
                        case 1:
                            buffer = _a.sent();
                            a = document.createElement('a');
                            fileUrl = URL.createObjectURL(new Blob([buffer]));
                            a.href = fileUrl;
                            a.download = filename + ".xlsx";
                            a.click();
                            URL.revokeObjectURL(fileUrl);
                            return [2 /*return*/];
                    }
                });
            });
        },
        workbook: workbook,
    };
};
exports.createExcel = createExcel;
