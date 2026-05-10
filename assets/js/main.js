/* NEVAGS Eco Brick & Construction — Main JS */

// ── Nav scroll effect ──────────────────────────────────────
const nav = document.getElementById('nav');
window.addEventListener('scroll', () => {
  nav.classList.toggle('scrolled', window.scrollY > 60);
});

// ── Mobile menu ────────────────────────────────────────────
document.getElementById('hamburger').addEventListener('click', () => {
  document.getElementById('mob-nav').classList.toggle('open');
});
document.querySelectorAll('#mob-nav a').forEach(a => {
  a.addEventListener('click', () => document.getElementById('mob-nav').classList.remove('open'));
});

// ── Smooth scroll for all anchor links ────────────────────
document.querySelectorAll('a[href^="#"]').forEach(a => {
  a.addEventListener('click', e => {
    const target = document.querySelector(a.getAttribute('href'));
    if (target) {
      e.preventDefault();
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  });
});

// ── Scroll spy (active nav link) ──────────────────────────
const sections = document.querySelectorAll('section[id], div[id="stats"]');
const navLinks = document.querySelectorAll('.nl');
const spy = new IntersectionObserver(entries => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      navLinks.forEach(l => l.classList.remove('active'));
      const match = document.querySelector(`.nl[href="#${e.target.id}"]`);
      if (match) match.classList.add('active');
    }
  });
}, { rootMargin: '-50% 0px -50% 0px' });
sections.forEach(s => spy.observe(s));

// ── Fade-up on scroll ─────────────────────────────────────
const fadeObs = new IntersectionObserver(entries => {
  entries.forEach(e => { if (e.isIntersecting) e.target.classList.add('vis'); });
}, { threshold: 0.12 });
document.querySelectorAll('.fu').forEach(el => fadeObs.observe(el));

// ── Animated counters ─────────────────────────────────────
function animateCounter(el, target, suffix) {
  const isFloat = target % 1 !== 0;
  const duration = 1800;
  const start = performance.now();
  const step = ts => {
    const progress = Math.min((ts - start) / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3);
    const val = ease * target;
    el.textContent = (isFloat ? val.toFixed(2) : Math.floor(val)).toLocaleString() + (suffix || '');
    if (progress < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}
const counterObs = new IntersectionObserver(entries => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      const el = e.target;
      animateCounter(el, parseFloat(el.dataset.target), el.dataset.suffix || '');
      counterObs.unobserve(el);
    }
  });
}, { threshold: 0.5 });
document.querySelectorAll('.counter[data-target]').forEach(el => counterObs.observe(el));

// ── Progress bars ─────────────────────────────────────────
const barObs = new IntersectionObserver(entries => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      e.target.querySelectorAll('.pfill[data-width]').forEach(bar => {
        setTimeout(() => { bar.style.width = bar.dataset.width + '%'; }, 200);
      });
      barObs.unobserve(e.target);
    }
  });
}, { threshold: 0.3 });
document.querySelectorAll('section, .card').forEach(el => {
  if (el.querySelector('.pfill')) barObs.observe(el);
});

// ── Chart.js defaults ─────────────────────────────────────
Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
Chart.defaults.color = '#64748b';

const FOREST  = '#1B4332';
const FMID    = '#40916C';
const FPALE   = '#D8F3DC';
const BRAND   = '#E8690A';
const BRAND2  = '#F59E0B';
const CHAR    = '#1C2B28';

// ── Revenue & Profit Chart ────────────────────────────────
new Chart(document.getElementById('revenueChart'), {
  type: 'bar',
  data: {
    labels: ['Ordinary Bricks', 'Face Bricks (Premium)'],
    datasets: [
      {
        label: 'Annual Revenue (MWK M)',
        data: [900, 1170],
        backgroundColor: [FOREST, BRAND],
        borderRadius: 8,
        borderSkipped: false,
      },
      {
        label: 'Annual Profit (MWK M)',
        data: [360, 414],
        backgroundColor: [FMID, BRAND2],
        borderRadius: 8,
        borderSkipped: false,
      }
    ]
  },
  options: {
    responsive: true,
    plugins: {
      legend: { position: 'top', labels: { boxWidth: 12, padding: 16 } },
      tooltip: {
        callbacks: {
          label: ctx => ` K${ctx.parsed.y.toLocaleString()}M`
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        grid: { color: '#f1f5f9' },
        ticks: { callback: v => `K${v}M` }
      },
      x: { grid: { display: false } }
    }
  }
});

// ── Revenue Growth Projection ─────────────────────────────
new Chart(document.getElementById('growthChart'), {
  type: 'line',
  data: {
    labels: ['Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'],
    datasets: [
      {
        label: 'Revenue (USD)',
        data: [600000, 1200000, 2500000, 3200000, 4100000],
        borderColor: FOREST,
        backgroundColor: 'rgba(27,67,50,0.08)',
        fill: true,
        tension: 0.4,
        pointBackgroundColor: FOREST,
        pointRadius: 5,
      },
      {
        label: 'Gross Profit (USD)',
        data: [220000, 440000, 900000, 1200000, 1600000],
        borderColor: BRAND,
        backgroundColor: 'rgba(232,105,10,0.06)',
        fill: true,
        tension: 0.4,
        pointBackgroundColor: BRAND,
        pointRadius: 5,
      }
    ]
  },
  options: {
    responsive: true,
    plugins: {
      legend: { position: 'top', labels: { boxWidth: 12, padding: 16 } },
      tooltip: {
        callbacks: {
          label: ctx => ` $${ctx.parsed.y.toLocaleString()}`
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        grid: { color: '#f1f5f9' },
        ticks: { callback: v => `$${(v/1000).toFixed(0)}K` }
      },
      x: { grid: { display: false } }
    }
  }
});

// ── Gender Donut Chart ────────────────────────────────────
new Chart(document.getElementById('genderChart'), {
  type: 'doughnut',
  data: {
    labels: ['Male (36)', 'Female (15)'],
    datasets: [{
      data: [36, 15],
      backgroundColor: [FOREST, BRAND],
      borderColor: ['#fff', '#fff'],
      borderWidth: 3,
      hoverOffset: 8,
    }]
  },
  options: {
    cutout: '68%',
    plugins: {
      legend: { position: 'bottom', labels: { boxWidth: 12, padding: 16 } },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/51*100)}%)`
        }
      }
    }
  }
});

// ── Inclusion bar chart ───────────────────────────────────
new Chart(document.getElementById('inclusionChart'), {
  type: 'bar',
  data: {
    labels: ['Current\n(2026)', 'Phase 3\nTarget', 'Full Scale\nTarget'],
    datasets: [
      {
        label: 'Female',
        data: [15, 18, 25],
        backgroundColor: BRAND,
        borderRadius: 6,
        borderSkipped: false,
      },
      {
        label: 'Male',
        data: [36, 27, 50],
        backgroundColor: FOREST,
        borderRadius: 6,
        borderSkipped: false,
      }
    ]
  },
  options: {
    responsive: true,
    indexAxis: 'y',
    plugins: {
      legend: { position: 'top', labels: { boxWidth: 12, padding: 16 } },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.x} staff`
        }
      }
    },
    scales: {
      x: {
        stacked: true,
        grid: { color: '#f1f5f9' },
        ticks: { callback: v => `${v}` }
      },
      y: { stacked: true, grid: { display: false } }
    }
  }
});

// ── Staffing Trajectory Chart ─────────────────────────────
new Chart(document.getElementById('staffChart'), {
  type: 'line',
  data: {
    labels: ['Wk1','Wk2','Wk3','Wk4','Wk5','Wk6','Wk7','Wk8','Wk9','Wk10','Wk11','Wk12','Wk13'],
    datasets: [{
      label: 'Staff on Site',
      data: [18, 18, 33, 33, 39, 39, 40, 40, 45, 45, 45, 45, 45],
      borderColor: FOREST,
      backgroundColor: 'rgba(27,67,50,0.1)',
      fill: true,
      tension: 0.3,
      pointBackgroundColor: ctx => {
        const vals = [18,18,33,33,39,39,40,40,45,45,45,45,45];
        const prev = ctx.dataIndex > 0 ? vals[ctx.dataIndex-1] : vals[0];
        return vals[ctx.dataIndex] > prev ? BRAND : FOREST;
      },
      pointRadius: 5,
      stepped: 'after',
    }]
  },
  options: {
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` ${ctx.parsed.y} staff on site`
        }
      }
    },
    scales: {
      y: {
        beginAtZero: false,
        min: 10,
        grid: { color: '#f1f5f9' },
        ticks: { stepSize: 5, callback: v => `${v} staff` }
      },
      x: { grid: { display: false } }
    }
  }
});

// ── Cost Per m² Comparison Chart ─────────────────────────
const cmpCanvas = document.getElementById('costCompareChart');
if (cmpCanvas) {
  new Chart(cmpCanvas, {
    type: 'bar',
    data: {
      labels: ['Cement Blocks\n(12 blocks/m²)', 'NEVAGS VSK Bricks\n(59 bricks/m²)', 'Traditional Firewood Bricks\n(NOW ILLEGAL)'],
      datasets: [{
        label: 'Material Cost per m² (MWK)',
        data: [36000, 17700, 8850],
        backgroundColor: ['#DC2626', '#1B4332', '#D97706'],
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` K${ctx.parsed.y.toLocaleString()} per m²`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grid: { color: '#f1f5f9' },
          ticks: { callback: v => `K${(v/1000).toFixed(0)}K` }
        },
        x: {
          grid: { display: false },
          ticks: { maxRotation: 0, font: { size: 10 } }
        }
      }
    }
  });
}

// ── Budget Horizontal Bar Chart ───────────────────────────
new Chart(document.getElementById('budgetChart'), {
  type: 'bar',
  data: {
    labels: [
      'Salaries (3 mo)',
      'Raw Materials',
      'Machinery & Tools',
      'Office & IT',
      'PPE & Safety',
      'Utilities & Energy',
      'Marketing',
      'Insurance & Admin',
    ],
    datasets: [{
      label: 'MWK (Millions)',
      data: [20.3, 7.5, 4.6, 2.6, 4.7, 1.6, 1.0, 0.7],
      backgroundColor: [FOREST, BRAND, FMID, BRAND2, FOREST, FMID, BRAND, CHAR],
      borderRadius: 6,
      borderSkipped: false,
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ` K${ctx.parsed.x.toFixed(1)}M`
        }
      }
    },
    scales: {
      x: {
        beginAtZero: true,
        grid: { color: '#f1f5f9' },
        ticks: { callback: v => `K${v}M` }
      },
      y: { grid: { display: false } }
    }
  }
});
