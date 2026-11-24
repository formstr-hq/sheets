// Stand-in for `readable-stream` in the browser bundle. exceljs/lib/utils/parse-sax requires
// PassThrough to upgrade legacy Node streams; we only feed it async iterables, so it never runs.
module.exports = {
  PassThrough: class PassThrough {
    constructor() {
      throw new Error("Node streams are not supported in the browser import");
    }
  },
};
