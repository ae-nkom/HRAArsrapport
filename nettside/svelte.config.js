import adapter from "@sveltejs/adapter-static";

/** @type {import("@sveltejs/kit").Config} */
const config = {
  kit: {
    adapter: adapter(),
    paths: {
      base: process.env.BASE_PATH || ""
    },
    csp: {
      mode: "hash",
      directives: {
        "default-src": ["self"],
        "script-src": ["self"],
        "style-src": ["self", "unsafe-inline"],
        "img-src": ["self", "data:", "blob:"],
        "font-src": ["self", "data:"],
        "connect-src": process.env.NODE_ENV === "production" ? ["none"] : ["self", "ws:", "wss:"],
        "worker-src": ["self"],
        "object-src": ["none"],
        "base-uri": ["self"],
        "form-action": ["none"]
      }
    },
    files: {
      lib: "src/components"
    }
  }
};

export default config;
