import adapter from "@sveltejs/adapter-static";

/** @type {import("@sveltejs/kit").Config} */
const config = {
  kit: {
    adapter: adapter(),
    csp: {
      mode: "hash",
      directives: {
        "default-src": ["self"],
        "script-src": ["self"],
        "style-src": ["self", "unsafe-inline"],
        "img-src": ["self", "data:", "blob:"],
        "font-src": ["self", "data:"],
        "connect-src": ["self", "ws:", "wss:"],
        "object-src": ["none"],
        "base-uri": ["self"],
        "form-action": ["self"]
      }
    },
    files: {
      lib: "src/components"
    }
  }
};

export default config;
