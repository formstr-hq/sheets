// Stand-in for Node's `fs` in the browser bundle. ExcelJS's workbook reader requires it, but the
// streaming import only uses ExcelJS's worksheet/styles parsers, which never touch it.
module.exports = {};
