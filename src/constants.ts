export const CATEGORY_COLORS = [
  { name: "Red", value: "#b10e1c", preset: "Preset0" },
  { name: "Orange", value: "#c33400", preset: "Preset1" },
  { name: "Peach", value: "#e69a3e", preset: "Preset2" },
  { name: "Yellow", value: "#e3cc00", preset: "Preset3" },
  { name: "Light Green", value: "#009c4e", preset: "Preset4" },
  { name: "Light Teal", value: "#00a3ae", preset: "Preset5" },
  { name: "Lime Green", value: "#a8cc4d", preset: "Preset6" },
  { name: "Blue", value: "#006cbe", preset: "Preset7" },
  { name: "Lavender", value: "#756cc8", preset: "Preset8" },
  { name: "Magenta", value: "#cc007e", preset: "Preset9" },
  { name: "Light Gray", value: "#919da1", preset: "Preset10" },
  { name: "Steel", value: "#005265", preset: "Preset11" },
  { name: "Warm Gray", value: "#8c8e83", preset: "Preset12" },
  { name: "Gray", value: "#5d6c70", preset: "Preset13" },
  { name: "Dark Gray", value: "#3e3e3e", preset: "Preset14" },
  { name: "Dark Red", value: "#6a0a1a", preset: "Preset15" },
  { name: "Dark Orange", value: "#b5490f", preset: "Preset16" },
  { name: "Brown", value: "#814e29", preset: "Preset17" },
  { name: "Gold", value: "#ae8e00", preset: "Preset18" },
  { name: "Dark Green", value: "#0a600a", preset: "Preset19" },
  { name: "Teal", value: "#02767a", preset: "Preset20" },
  { name: "Green", value: "#427505", preset: "Preset21" },
  { name: "Navy Blue", value: "#00345c", preset: "Preset22" },
  { name: "Dark Purple", value: "#684697", preset: "Preset23" },
  { name: "Dark Pink", value: "#8c0059", preset: "Preset24" },
];

// Are we running in Internet Explorer mode?
export const IS_IE: boolean = /MSIE|Trident/.test(window.navigator.userAgent);

export const DEFAULT_ADDIN_CATEGORIES = {
  generalCategory: {
    displayName: "Mail Notes",
    color: "Preset7",
  },
  messageCategory: {
    displayName: "Message - Mail Notes",
    color: "Preset6",
  },
  conversationCategory: {
    displayName: "Conversation - Mail Notes",
    color: "Preset5",
  },
  senderCategory: {
    displayName: "Sender - Mail Notes",
    color: "Preset8",
  },
};
