(function () {
  const THEME_KEY = 'taskTrackerTheme';
  const root = document.documentElement;

  function systemPrefersDark() {
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  function effectiveTheme() {
    const saved = localStorage.getItem(THEME_KEY);
    return saved === 'light' || saved === 'dark' ? saved : (systemPrefersDark() ? 'dark' : 'light');
  }

  function applySavedTheme() {
    const saved = localStorage.getItem(THEME_KEY);
    if (saved === 'light' || saved === 'dark') {
      root.setAttribute('data-theme', saved);
    } else {
      root.removeAttribute('data-theme');
    }
  }

  // Runs immediately (this script is loaded in <head>, before the body
  // paints) so there's no flash of the wrong theme on load.
  applySavedTheme();

  window.initThemeToggle = function (buttonId) {
    const btn = document.getElementById(buttonId);
    if (!btn) return;

    function updateIcon() {
      btn.textContent = effectiveTheme() === 'dark' ? '☀️' : '🌙';
    }

    updateIcon();
    btn.addEventListener('click', () => {
      const next = effectiveTheme() === 'dark' ? 'light' : 'dark';
      localStorage.setItem(THEME_KEY, next);
      applySavedTheme();
      updateIcon();
    });
  };
})();
