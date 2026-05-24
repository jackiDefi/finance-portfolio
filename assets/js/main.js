(function () {
  'use strict';

  const STORAGE_KEY = 'portfolio-theme';
  const root = document.documentElement;

  function apply(theme) {
    root.setAttribute('data-theme', theme);
    const btn = document.querySelector('.theme-toggle');
    if (btn) btn.textContent = theme === 'dark' ? 'Light' : 'Dark';
  }

  // Apply theme as early as possible to avoid flash.
  const saved = (function () { try { return localStorage.getItem(STORAGE_KEY); } catch (e) { return null; } })();
  const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  apply(saved || (prefersDark ? 'dark' : 'light'));

  document.addEventListener('DOMContentLoaded', function () {
    const btn = document.querySelector('.theme-toggle');
    if (!btn) return;
    btn.textContent = root.getAttribute('data-theme') === 'dark' ? 'Light' : 'Dark';
    btn.addEventListener('click', function () {
      const current = root.getAttribute('data-theme') || 'light';
      const next = current === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem(STORAGE_KEY, next); } catch (e) {}
      apply(next);
    });
  });
})();
