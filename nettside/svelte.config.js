import preprocess from "svelte-preprocess";
import adapter from "@sveltejs/adapter-static";

/** @type {import("@sveltejs/kit").Config} */
const config = {
  preprocess: preprocess({
    postcss: true
  }),
  kit: {
    adapter: adapter(),
    files: {
      lib: "src/components"
    }
  }
};

export default config;
