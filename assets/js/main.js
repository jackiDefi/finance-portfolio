(function () {
  'use strict';

  const STORAGE_KEY = 'portfolio-theme';
  const root = document.documentElement;

  function apply(theme) {
    // Dark is the default (no data-theme attribute = dark via :root).
    // Only set the attribute when explicitly switching to light.
    if (theme === 'light') {
      root.setAttribute('data-theme', 'light');
    } else {
      root.removeAttribute('data-theme');
    }
    const btn = document.querySelector('.theme-toggle');
    if (btn) btn.textContent = theme === 'light' ? 'Dark' : 'Light';
  }

  // Read persisted preference; default to dark.
  const saved = (function () {
    try { return localStorage.getItem(STORAGE_KEY); } catch (e) { return null; }
  })();
  apply(saved === 'light' ? 'light' : 'dark');

  document.addEventListener('DOMContentLoaded', function () {
    const btn = document.querySelector('.theme-toggle');
    if (!btn) return;
    const current = root.getAttribute('data-theme') === 'light' ? 'light' : 'dark';
    btn.textContent = current === 'light' ? 'Dark' : 'Light';
    btn.addEventListener('click', function () {
      const cur = root.getAttribute('data-theme') === 'light' ? 'light' : 'dark';
      const next = cur === 'light' ? 'dark' : 'light';
      try { localStorage.setItem(STORAGE_KEY, next); } catch (e) {}
      apply(next);
      // Notify any chart canvases that the palette has changed.
      window.dispatchEvent(new CustomEvent('themechange', { detail: { theme: next } }));
    });
  });

  // ---------------------------------------------------------------
  // Hero canvas — animated multi-line "sector index" chart drawn on
  // page load. Five lines representing the five portfolio industries.
  // ---------------------------------------------------------------
  function renderHeroChart() {
    const canvas = document.getElementById('hero-chart');
    if (!canvas) return;

    function resize() {
      const dpr = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      canvas.width = Math.max(1, Math.floor(rect.width * dpr));
      canvas.height = Math.max(1, Math.floor(rect.height * dpr));
      canvas.style.width = rect.width + 'px';
      canvas.style.height = rect.height + 'px';
      const ctx = canvas.getContext('2d');
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      return { ctx: ctx, w: rect.width, h: rect.height };
    }

    function colors() {
      const s = getComputedStyle(document.documentElement);
      return {
        accent: s.getPropertyValue('--accent').trim(),
        bright: s.getPropertyValue('--accent-bright').trim(),
        muted:  s.getPropertyValue('--accent-muted').trim(),
        secondary: s.getPropertyValue('--accent-secondary').trim(),
        positive: s.getPropertyValue('--positive').trim(),
        grid:   s.getPropertyValue('--grid-line').trim(),
      };
    }

    // Generate stylised industry series with light noise + persistent trend.
    function makeSeries(seed, points, trend, volatility) {
      let rand = (function (s) {
        return function () { s = (s * 9301 + 49297) % 233280; return s / 233280; };
      })(seed);
      const data = [];
      let v = 0.45 + rand() * 0.1;
      for (let i = 0; i < points; i++) {
        const noise = (rand() - 0.5) * volatility;
        const drift = trend * 0.012;
        v += noise + drift;
        v = Math.max(0.05, Math.min(0.95, v));
        data.push(v);
      }
      return data;
    }

    const POINTS = 80;
    const series = [
      { name: 'Telecom',   data: makeSeries(7,  POINTS, 1.0, 0.04), key: 'accent',    width: 1.8 },
      { name: 'Oil&Gas',   data: makeSeries(13, POINTS, 0.4, 0.06), key: 'bright',    width: 1.4 },
      { name: 'Banking',   data: makeSeries(29, POINTS, 0.7, 0.03), key: 'secondary', width: 1.4 },
      { name: 'Agri',      data: makeSeries(41, POINTS, 0.3, 0.05), key: 'positive',  width: 1.2 },
      { name: 'Retail',    data: makeSeries(67, POINTS, 0.6, 0.035),key: 'muted',     width: 1.2 },
    ];

    let progress = 0;            // 0 → 1 animation
    let animStart = null;
    let lastDrawn = -1;
    const DURATION = 1800;       // ms — line draw animation

    function draw(p) {
      const { ctx, w, h } = resize();
      const c = colors();
      ctx.clearRect(0, 0, w, h);

      // --- subtle grid (horizontal lines) ---
      ctx.strokeStyle = c.grid;
      ctx.lineWidth = 1;
      for (let i = 1; i < 6; i++) {
        const y = (h / 6) * i;
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(w, y);
        ctx.stroke();
      }
      // Vertical ticks (right-edge style)
      ctx.strokeStyle = c.grid;
      for (let i = 1; i < 10; i++) {
        const x = (w / 10) * i;
        ctx.beginPath();
        ctx.moveTo(x, h - 6);
        ctx.lineTo(x, h);
        ctx.stroke();
      }

      // --- sector lines ---
      const visiblePoints = Math.max(2, Math.floor(p * series[0].data.length));
      series.forEach(function (s) {
        ctx.strokeStyle = c[s.key] || c.accent;
        ctx.lineWidth = s.width;
        ctx.lineJoin = 'round';
        ctx.lineCap = 'round';
        ctx.globalAlpha = 0.85;
        ctx.beginPath();
        for (let i = 0; i < visiblePoints; i++) {
          const x = (w / (s.data.length - 1)) * i;
          const y = h - (s.data[i] * h * 0.78) - (h * 0.06);
          if (i === 0) ctx.moveTo(x, y);
          else ctx.lineTo(x, y);
        }
        ctx.stroke();
        ctx.globalAlpha = 1;
      });

      // --- subtle area fill below leading line ---
      const lead = series[0];
      const grad = ctx.createLinearGradient(0, 0, 0, h);
      grad.addColorStop(0, c.accent + '22');
      grad.addColorStop(1, c.accent + '00');
      ctx.fillStyle = grad;
      ctx.beginPath();
      for (let i = 0; i < visiblePoints; i++) {
        const x = (w / (lead.data.length - 1)) * i;
        const y = h - (lead.data[i] * h * 0.78) - (h * 0.06);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      const lastX = (w / (lead.data.length - 1)) * (visiblePoints - 1);
      ctx.lineTo(lastX, h);
      ctx.lineTo(0, h);
      ctx.closePath();
      ctx.fill();
    }

    function step(ts) {
      if (animStart === null) animStart = ts;
      progress = Math.min(1, (ts - animStart) / DURATION);
      // ease-out cubic
      const eased = 1 - Math.pow(1 - progress, 3);
      draw(eased);
      if (progress < 1) requestAnimationFrame(step);
    }

    function redrawStatic() {
      draw(1);
    }

    // Initial animated draw
    requestAnimationFrame(step);

    // Redraw on theme change & resize
    window.addEventListener('resize', redrawStatic);
    window.addEventListener('themechange', redrawStatic);
  }

  document.addEventListener('DOMContentLoaded', renderHeroChart);
})();
