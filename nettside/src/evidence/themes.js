const sharedColors = {
  primary: "#0f2747",
  accent: "#214e86",
  "base-100": "#ffffff",
  "base-200": "#f2f6fb",
  "base-300": "#d6e0ec",
  "base-content": "#10233f",
  "base-content-muted": "#5e718d",
  success: "#15803d",
  warning: "#b45309",
  error: "#b91c1c"
};

export const themes = {
  light: {
    colors: sharedColors,
    colorPalettes: {
      default: ["#0f766e", "#c2410c", "#2563eb", "#ca8a04", "#7c3aed"]
    },
    colorScales: {
      default: ["#f8fafc", "#0f2747"]
    }
  },
  dark: {
    colors: {
      ...sharedColors,
      primary: "#9fc2eb",
      accent: "#79a9df",
      "base-100": "#09121f",
      "base-200": "#0f1d31",
      "base-300": "#1c314f",
      "base-content": "#eef4fb",
      "base-content-muted": "#b7c8dc"
    },
    colorPalettes: {
      default: ["#5eead4", "#fdba74", "#93c5fd", "#fde68a", "#c4b5fd"]
    },
    colorScales: {
      default: ["#7aa4d6", "#eef4fb"]
    }
  }
};

export const themesConfig = {
  appearance: {
    default: "light",
    switcher: false
  }
};
