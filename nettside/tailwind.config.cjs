const evidenceTailwind = require("@evidence-dev/tailwind/config").config;

/** @type {import("tailwindcss").Config} */
module.exports = {
  content: ["./src/**/*.{html,js,svelte,ts}"],
  presets: [evidenceTailwind],
  theme: {
    extend: {}
  },
  plugins: []
};
