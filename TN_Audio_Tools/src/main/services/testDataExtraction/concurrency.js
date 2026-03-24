function clampMaxConcurrentTasks(value, fallbackValue = 1) {
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    return Math.max(1, Math.trunc(Number(fallbackValue) || 1));
  }

  return Math.max(1, Math.trunc(numericValue));
}

module.exports = {
  clampMaxConcurrentTasks
};