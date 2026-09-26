// @opendocument/odr-core is an Emscripten build that names Node's `module`
// in a branch the browser never takes; webpack must not try to bundle it
module.exports = {
  resolve: {
    fallback: { module: false }
  }
};
