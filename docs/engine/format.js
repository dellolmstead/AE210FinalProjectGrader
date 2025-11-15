const SPECIFIER = /%(\d+)?(?:\.(\d+))?([dfs])/g;
const PERCENT_SENTINEL = "\uFFFF";

function formatNumber(value, decimals) {
  const num = Number(value);
  if (!Number.isFinite(num)) {
    return "NaN";
  }
  if (decimals != null) {
    return num.toFixed(decimals);
  }
  return num.toString();
}

export function format(template, ...args) {
  let index = 0;
  const sanitized = template.replace(/%%/g, PERCENT_SENTINEL);
  const formatted = sanitized.replace(SPECIFIER, (_, width, precision, type) => {
    const value = args[index++];
    if (type === "s") {
      return String(value);
    }
    if (type === "d") {
      const num = Number(value);
      return Number.isFinite(num) ? Math.trunc(num).toString() : "NaN";
    }
    if (type === "f") {
      const decimals = precision != null ? Number(precision) : undefined;
      return formatNumber(value, decimals);
    }
    return "";
  });
  return formatted.replace(new RegExp(PERCENT_SENTINEL, "g"), "%");
}
