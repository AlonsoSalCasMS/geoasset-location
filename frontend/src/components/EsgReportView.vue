<template>
  <div class="report-page">

    <!-- Top bar -->
    <div class="report-topbar">
      <div class="report-topbar-left">
        <button class="back-btn" @click="store.setView('esg')">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
        </button>
        <div>
          <div class="report-overline">Informe Ejecutivo ESG</div>
          <div class="report-company">{{ store.selectedCompany?.name || store.metadata?.company?.name }}</div>
        </div>
      </div>
      <div class="report-topbar-right">
        <span class="report-date-chip">{{ today }}</span>
        <button v-if="store.esgReport" class="print-btn" @click="openPrintWindow">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <polyline points="6 9 6 2 18 2 18 9" />
            <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
            <rect x="6" y="14" width="12" height="8" />
          </svg>
          Exportar PDF
        </button>
      </div>
    </div>

    <!-- Loading -->
    <section v-if="loading" class="report-fullscreen">
      <div class="hero-noise--report" />
      <div class="report-loading-inner">
        <div class="report-spinner">
          <svg width="56" height="56" viewBox="0 0 56 56">
            <circle cx="28" cy="28" r="24" fill="none" stroke="rgba(21,101,192,0.12)" stroke-width="3" />
            <circle cx="28" cy="28" r="24" fill="none" stroke="url(#repGrad)" stroke-width="3"
              stroke-dasharray="40 110" stroke-linecap="round" class="report-spin-arc" />
            <defs>
              <linearGradient id="repGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stop-color="#1565c0" stop-opacity="0" />
                <stop offset="100%" stop-color="#1565c0" />
              </linearGradient>
            </defs>
          </svg>
        </div>
        <h3 class="report-loading-title">Generando informe ejecutivo</h3>
        <p class="report-loading-sub">El modelo está redactando el análisis ESG...</p>
        <button class="cancel-link" @click="store.setView('esg')">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
          Cancelar
        </button>
      </div>
    </section>

    <!-- Error -->
    <section v-else-if="error" class="report-fullscreen">
      <div class="hero-noise--report" />
      <div class="report-loading-inner">
        <svg width="52" height="52" viewBox="0 0 52 52" style="margin-bottom:1.5rem">
          <circle cx="26" cy="26" r="24" fill="rgba(239,68,68,0.08)" stroke="#ef4444" stroke-width="1.5" />
          <path d="M26 16v12M26 32v2" stroke="#ef4444" stroke-width="2.5" stroke-linecap="round" />
        </svg>
        <h3 class="report-loading-title" style="color:#b91c1c">Error generando el informe</h3>
        <p class="report-loading-sub">{{ error }}</p>
        <button class="retry-btn" @click="generateReport">Reintentar</button>
        <button class="cancel-link" @click="store.setView('esg')">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
          Volver al análisis
        </button>
      </div>
    </section>

    <!-- Report content -->
    <div v-else-if="store.esgReport" class="report-body">

      <!-- Cover strip -->
      <div class="report-cover">
        <div class="report-cover-left">
          <div class="report-cover-label">INFORME EJECUTIVO ESG</div>
          <h1 class="report-cover-title">{{ store.esgReport.report_title }}</h1>
          <div class="report-cover-meta">
            <span>{{ store.selectedCompany?.name || store.metadata?.company?.name }}</span>
            <span class="meta-sep">·</span>
            <span>{{ store.assets.length }} activos analizados</span>
            <span class="meta-sep">·</span>
            <span>{{ today }}</span>
          </div>
        </div>
        <div v-if="store.esgAnalysis" class="report-cover-score">
          <div class="cover-grade" :class="`grade-${store.esgAnalysis.grade}`">
            {{ store.esgAnalysis.grade }}
          </div>
          <div class="cover-score-val">{{ store.esgAnalysis.overall_score }}<span>/100</span></div>
          <div class="cover-score-lbl">ESG Score</div>
        </div>
      </div>

      <!-- Score snapshot strip -->
      <div v-if="store.esgAnalysis" class="score-strip">
        <div v-for="dim in scoreDims" :key="dim.key" class="score-strip-item">
          <div class="score-strip-val" :style="{ color: dim.color }">{{ dim.score }}</div>
          <div class="score-strip-label">{{ dim.label }}</div>
          <div class="score-strip-bar">
            <div class="score-strip-fill" :style="{ width: dim.score + '%', background: dim.color }" />
          </div>
        </div>
      </div>

      <!-- Sections -->
      <div class="report-sections">

        <!-- Executive summary -->
        <div class="report-section">
          <div class="section-eyebrow">01</div>
          <h2 class="section-title">Resumen Ejecutivo</h2>
          <div class="section-body">
            <p v-for="(para, i) in execParas" :key="i" class="report-para">{{ para }}</p>
          </div>
        </div>

        <!-- Key findings -->
        <div class="report-section">
          <div class="section-eyebrow">02</div>
          <h2 class="section-title">Hallazgos Principales</h2>
          <div class="section-body">
            <div v-for="(f, i) in store.esgReport.key_findings" :key="i" class="finding-row">
              <div class="finding-num">{{ (i + 1).toString().padStart(2, '0') }}</div>
              <p class="finding-text">{{ f }}</p>
            </div>
          </div>
        </div>

        <!-- Top risks -->
        <div class="report-section">
          <div class="section-eyebrow">03</div>
          <h2 class="section-title">Riesgos Prioritarios</h2>
          <div class="section-body">
            <div v-for="risk in store.esgReport.top_risks" :key="risk.rank" class="risk-card">
              <div class="risk-card-header">
                <div class="risk-rank">{{ risk.rank }}</div>
                <h3 class="risk-title">{{ risk.title }}</h3>
              </div>
              <p class="risk-desc">{{ risk.description }}</p>
              <div class="risk-mitigation">
                <div class="risk-mit-label">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
                    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                  </svg>
                  Mitigación recomendada
                </div>
                <p class="risk-mit-text">{{ risk.mitigation }}</p>
              </div>
            </div>
          </div>
        </div>

        <!-- Action plan -->
        <div class="report-section">
          <div class="section-eyebrow">04</div>
          <h2 class="section-title">Plan de Acción</h2>
          <div class="section-body">
            <div class="action-table-wrap">
              <table class="action-table">
                <thead>
                  <tr>
                    <th>Prioridad</th>
                    <th>Área</th>
                    <th>Acción recomendada</th>
                    <th>Plazo</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(rec, i) in store.esgReport.recommendations" :key="i">
                    <td>
                      <span class="priority-badge" :class="`priority--${rec.priority.toLowerCase()}`">
                        {{ rec.priority }}
                      </span>
                    </td>
                    <td>
                      <span class="area-badge" :class="`area--${rec.area.toLowerCase()}`">
                        {{ rec.area.replace('_', ' ') }}
                      </span>
                    </td>
                    <td class="action-text">{{ rec.action }}</td>
                    <td class="timeline-text">{{ rec.timeline }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Regulatory roadmap -->
        <div class="report-section">
          <div class="section-eyebrow">05</div>
          <h2 class="section-title">Hoja de Ruta Regulatoria</h2>
          <div class="section-body">
            <p class="report-para">{{ store.esgReport.regulatory_roadmap }}</p>
          </div>
        </div>

        <!-- Conclusion -->
        <div class="report-section report-section--last">
          <div class="section-eyebrow">06</div>
          <h2 class="section-title">Conclusión</h2>
          <div class="section-body">
            <p class="report-para report-para--highlight">{{ store.esgReport.conclusion }}</p>
          </div>
        </div>

      </div>

      <!-- Footer -->
      <div class="report-footer">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
          <circle cx="12" cy="12" r="10" />
          <path d="M12 8v4l3 3" />
        </svg>
        Generado por GeoAssets Intelligence · {{ today }}
      </div>
    </div>

  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from 'vue';
import { useAppStore } from '@/stores/store';
import { generateEsgReport } from '@/services/backend';

const store = useAppStore();
const loading = ref(false);
const error = ref<string | null>(null);

const today = new Date().toLocaleDateString('es-ES', { year: 'numeric', month: 'long', day: 'numeric' });

const generateReport = async () => {
  if (!store.esgAnalysis) { store.setView('esg'); return; }
  loading.value = true;
  error.value = null;
  try {
    const companyName = store.selectedCompany?.name || store.metadata?.company?.name || '';
    const result = await generateEsgReport(companyName, store.esgAnalysis);
    if (result && (result as any).error) throw new Error((result as any).error);
    store.setEsgReport(result);
  } catch (e: any) {
    error.value = e?.message || 'Error desconocido';
  } finally {
    loading.value = false;
  }
};

onMounted(() => { if (!store.esgReport) generateReport(); });

// ── Print: open standalone window ───────────────────────────────
const openPrintWindow = () => {
  const r = store.esgReport!;
  const analysis = store.esgAnalysis;
  const company = store.selectedCompany?.name || store.metadata?.company?.name || '';

  const dims = analysis ? [
    { label: 'Riesgo Físico (35%)', score: analysis.scores.physical_risk.score, color: '#c62828' },
    { label: 'Impacto Ambiental (25%)', score: analysis.scores.environmental_impact.score, color: '#2e7d32' },
    { label: 'Trans. Regulatoria (25%)', score: analysis.scores.regulatory_transition.score, color: '#1565c0' },
    { label: 'Resiliencia (15%)', score: analysis.scores.resilience.score, color: '#6a1b9a' },
  ] : [];

  const gradeColors: Record<string, string> = { A: '#2e7d32', B: '#558b2f', C: '#e65100', D: '#c62828', E: '#4a148c' };
  const gradeColor = analysis ? (gradeColors[analysis.grade] || '#666') : '#666';

  const priorityColors: Record<string, string> = { HIGH: '#c62828', MEDIUM: '#e65100', LOW: '#2e7d32' };
  const areaColors: Record<string, string> = { CSRD: '#1565c0', EU_TAXONOMY: '#2e7d32', PHYSICAL_RISK: '#e65100', SFDR: '#6a1b9a', GOVERNANCE: '#00897b' };

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Informe ESG – ${company}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;500;600;700;800&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Sora',sans-serif;color:#0d1f3c;background:white;font-size:13px;line-height:1.6}
  @page{margin:1.8cm 1.5cm}
  @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}

  /* Cover */
  .cover{background:linear-gradient(135deg,#1b5e20 0%,#2e7d32 45%,#0d47a1 100%);color:white;padding:2.5rem 2rem;display:flex;align-items:center;justify-content:space-between;margin-bottom:0;border-radius:0}
  .cover-left{}
  .cover-eyebrow{font-size:0.6rem;text-transform:uppercase;letter-spacing:0.14em;color:rgba(255,255,255,0.65);margin-bottom:0.5rem}
  .cover-title{font-size:1.4rem;font-weight:800;color:white;margin-bottom:0.6rem;line-height:1.25}
  .cover-meta{font-size:0.72rem;color:rgba(255,255,255,0.75);display:flex;gap:0.5rem;align-items:center;flex-wrap:wrap}
  .meta-sep{opacity:0.4}
  .cover-score-wrap{text-align:center;flex-shrink:0}
  .cover-grade{width:60px;height:60px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:1.8rem;font-weight:800;color:white;margin:0 auto 0.3rem;background:${gradeColor};box-shadow:0 4px 16px rgba(0,0,0,0.3)}
  .cover-score-num{font-size:1.5rem;font-weight:800;color:white}
  .cover-score-lbl{font-size:0.62rem;color:rgba(255,255,255,0.65)}

  /* Score strip */
  .score-strip{display:flex;background:#f8faff;border-bottom:1px solid #e8eef7;padding:0}
  .score-strip-item{flex:1;padding:0.7rem 1rem;border-right:1px solid #e8eef7}
  .score-strip-item:last-child{border-right:none}
  .score-val{font-size:1.4rem;font-weight:800}
  .score-lbl{font-size:0.62rem;color:#54668f;margin-bottom:0.3rem}
  .score-bar-track{height:4px;background:rgba(0,0,0,0.07);border-radius:2px;overflow:hidden}
  .score-bar-fill{height:100%;border-radius:2px}

  /* Sections */
  .sections{padding:1.5rem 2rem}
  .section{margin-bottom:2rem;page-break-inside:avoid}
  .section-eyebrow{font-size:0.6rem;font-weight:700;text-transform:uppercase;letter-spacing:0.12em;color:#9e9e9e;margin-bottom:0.2rem}
  .section-title{font-size:1rem;font-weight:800;color:#0d1f3c;margin-bottom:0.75rem;padding-bottom:0.5rem;border-bottom:2px solid #e8eef7}
  p{font-size:0.82rem;color:#2a3a5e;line-height:1.75;margin-bottom:0.5rem}

  /* Findings */
  .finding{display:flex;gap:0.75rem;align-items:flex-start;margin-bottom:0.6rem;padding:0.5rem 0.7rem;background:#f8faff;border-radius:6px;border-left:3px solid #1565c0}
  .finding-num{font-size:0.7rem;font-weight:800;color:#1565c0;min-width:22px}

  /* Risks */
  .risk-card{margin-bottom:1rem;border:1px solid #e8eef7;border-radius:8px;overflow:hidden}
  .risk-header{display:flex;align-items:center;gap:0.75rem;padding:0.6rem 0.9rem;background:#fdf2f2}
  .risk-rank-badge{width:26px;height:26px;border-radius:50%;background:#c62828;color:white;display:flex;align-items:center;justify-content:center;font-size:0.72rem;font-weight:800;flex-shrink:0}
  .risk-title{font-size:0.85rem;font-weight:700;color:#1a0000}
  .risk-desc{padding:0.5rem 0.9rem;font-size:0.78rem;color:#3a2020}
  .risk-mit{padding:0.5rem 0.9rem;background:#f0faf4;border-top:1px solid #e0f0e8}
  .risk-mit-lbl{font-size:0.65rem;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;color:#2e7d32;margin-bottom:2px}

  /* Action table */
  table{width:100%;border-collapse:collapse;font-size:0.78rem}
  th{background:#f0f4ff;padding:0.45rem 0.6rem;text-align:left;font-size:0.68rem;font-weight:700;color:#3a4a6b;border-bottom:2px solid #dde6f7}
  td{padding:0.4rem 0.6rem;border-bottom:1px solid #f0f4ff;vertical-align:top}
  .p-badge{padding:2px 7px;border-radius:4px;font-size:0.65rem;font-weight:700;color:white}
  .a-badge{padding:2px 7px;border-radius:4px;font-size:0.65rem;font-weight:600}

  /* Highlight para */
  .para-highlight{background:linear-gradient(135deg,rgba(21,101,192,0.04),rgba(46,125,50,0.04));border-left:3px solid #1565c0;padding:0.75rem 1rem;border-radius:0 6px 6px 0}

  /* Footer */
  .report-footer{text-align:center;font-size:0.65rem;color:#9e9e9e;padding:1.5rem 2rem;border-top:1px solid #e8eef7;margin-top:1rem}
</style>
</head>
<body>

<div class="cover">
  <div class="cover-left">
    <div class="cover-eyebrow">Informe Ejecutivo ESG · GeoAssets Intelligence</div>
    <div class="cover-title">${r.report_title}</div>
    <div class="cover-meta">
      <span>${company}</span><span class="meta-sep">·</span>
      <span>${store.assets.length} activos analizados</span><span class="meta-sep">·</span>
      <span>${today}</span>
    </div>
  </div>
  ${analysis ? `<div class="cover-score-wrap">
    <div class="cover-grade">${analysis.grade}</div>
    <div class="cover-score-num">${analysis.overall_score}<small style="font-size:0.8rem;opacity:0.7">/100</small></div>
    <div class="cover-score-lbl">ESG SCORE GLOBAL</div>
  </div>` : ''}
</div>

${analysis ? `<div class="score-strip">
  ${dims.map(d => `<div class="score-strip-item">
    <div class="score-val" style="color:${d.color}">${d.score}</div>
    <div class="score-lbl">${d.label}</div>
    <div class="score-bar-track"><div class="score-bar-fill" style="width:${d.score}%;background:${d.color}"></div></div>
  </div>`).join('')}
</div>` : ''}

<div class="sections">

  <div class="section">
    <div class="section-eyebrow">01</div>
    <div class="section-title">Resumen Ejecutivo</div>
    ${r.executive_summary.split('\n').filter((p: string) => p.trim()).map((p: string) => `<p>${p}</p>`).join('')}
  </div>

  <div class="section">
    <div class="section-eyebrow">02</div>
    <div class="section-title">Hallazgos Principales</div>
    ${r.key_findings.map((f: string, i: number) => `
      <div class="finding">
        <span class="finding-num">${String(i + 1).padStart(2, '0')}</span>
        <p style="margin:0">${f}</p>
      </div>`).join('')}
  </div>

  <div class="section">
    <div class="section-eyebrow">03</div>
    <div class="section-title">Riesgos Prioritarios</div>
    ${r.top_risks.map((risk: any) => `
      <div class="risk-card">
        <div class="risk-header">
          <div class="risk-rank-badge">${risk.rank}</div>
          <div class="risk-title">${risk.title}</div>
        </div>
        <p class="risk-desc">${risk.description}</p>
        <div class="risk-mit">
          <div class="risk-mit-lbl">Mitigación recomendada</div>
          <p style="margin:0;font-size:0.75rem;color:#1b5e20">${risk.mitigation}</p>
        </div>
      </div>`).join('')}
  </div>

  <div class="section">
    <div class="section-eyebrow">04</div>
    <div class="section-title">Plan de Acción</div>
    <table>
      <thead><tr><th>Prioridad</th><th>Área</th><th>Acción</th><th>Plazo</th></tr></thead>
      <tbody>
        ${r.recommendations.map((rec: any) => `
          <tr>
            <td><span class="p-badge" style="background:${priorityColors[rec.priority] || '#666'}">${rec.priority}</span></td>
            <td><span class="a-badge" style="background:${areaColors[rec.area] || '#ddd'}20;color:${areaColors[rec.area] || '#666'}">${rec.area.replace('_', ' ')}</span></td>
            <td>${rec.action}</td>
            <td style="white-space:nowrap;color:#54668f">${rec.timeline}</td>
          </tr>`).join('')}
      </tbody>
    </table>
  </div>

  <div class="section">
    <div class="section-eyebrow">05</div>
    <div class="section-title">Hoja de Ruta Regulatoria</div>
    <p>${r.regulatory_roadmap}</p>
  </div>

  <div class="section">
    <div class="section-eyebrow">06</div>
    <div class="section-title">Conclusión</div>
    <p class="para-highlight">${r.conclusion}</p>
  </div>

</div>

<div class="report-footer">Generado por GeoAssets Intelligence con IA · ${today}</div>

<script>window.onload=()=>{window.print()}<\/script>
</body>
</html>`;

  const win = window.open('', '_blank');
  if (win) {
    win.document.write(html);
    win.document.close();
  }
};

// ── Score dims ───────────────────────────────────────────────────
const scoreDims = computed(() => {
  if (!store.esgAnalysis) return [];
  const s = store.esgAnalysis.scores;
  return [
    { key: 'pr', label: 'Riesgo Físico (35%)', score: s.physical_risk.score, color: '#c62828' },
    { key: 'ei', label: 'Imp. Ambiental (25%)', score: s.environmental_impact.score, color: '#2e7d32' },
    { key: 'rt', label: 'Trans. Regulatoria (25%)', score: s.regulatory_transition.score, color: '#1565c0' },
    { key: 're', label: 'Resiliencia (15%)', score: s.resilience.score, color: '#6a1b9a' },
  ];
});

const execParas = computed(() =>
  (store.esgReport?.executive_summary || '').split('\n').filter((p) => p.trim())
);
</script>

<style scoped>
.report-page {
  min-height: 100vh;
  background: #f4f7fb;
  overflow-y: auto;
}

/* Topbar */
.report-topbar {
  background: linear-gradient(135deg, #0d47a1 0%, #1565c0 50%, #1b5e20 100%);
  padding: 0.85rem 1.5rem;
  display: flex; align-items: center; justify-content: space-between;
}
.report-topbar-left { display: flex; align-items: center; gap: 0.75rem; }
.report-topbar-right { display: flex; align-items: center; gap: 0.75rem; }

.back-btn {
  width: 32px; height: 32px; border-radius: 8px;
  background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2);
  color: white; display: flex; align-items: center; justify-content: center;
  cursor: pointer; transition: background 150ms; flex-shrink: 0;
}
.back-btn:hover { background: rgba(255,255,255,0.25); }

.report-overline { font-size: 0.62rem; text-transform: uppercase; letter-spacing: 0.1em; color: rgba(255,255,255,0.65); }
.report-company { font-size: 1.05rem; font-weight: 700; color: white; }
.report-date-chip { padding: 3px 10px; border-radius: 100px; border: 1px solid rgba(255,255,255,0.3); color: rgba(255,255,255,0.85); font-size: 0.72rem; }

.print-btn {
  display: inline-flex; align-items: center; gap: 0.4rem;
  background: rgba(255,255,255,0.9); color: #0d1f3c;
  border: none; border-radius: 8px; padding: 0.4rem 0.9rem;
  font-size: 0.78rem; font-weight: 600; cursor: pointer; transition: background 150ms;
}
.print-btn:hover { background: white; }

/* Fullscreen states */
.report-fullscreen {
  position: relative; min-height: calc(100vh - 68px);
  display: flex; align-items: center; justify-content: center; overflow: hidden;
}
.hero-noise--report {
  position: absolute; inset: 0;
  background:
    radial-gradient(circle at 20% 20%, rgba(21,101,192,0.1), transparent 36%),
    radial-gradient(circle at 80% 15%, rgba(46,125,50,0.08), transparent 36%),
    linear-gradient(160deg, rgba(255,255,255,0.9), rgba(240,245,255,0.95));
  pointer-events: none;
}
.report-loading-inner {
  position: relative; z-index: 1;
  display: flex; flex-direction: column; align-items: center; text-align: center;
  padding: 0 2rem;
}
.report-spinner { margin-bottom: 1.5rem; }
.report-spin-arc { transform-origin: center; animation: reportSpin 1.2s linear infinite; }
@keyframes reportSpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
.report-loading-title { font-size: 1.3rem; font-weight: 700; color: #0d1f3c; margin-bottom: 0.4rem; }
.report-loading-sub { font-size: 0.9rem; color: #54668f; margin-bottom: 1.5rem; }

.retry-btn {
  background: linear-gradient(135deg, #1565c0, #0d47a1); color: white;
  border: none; border-radius: 10px; padding: 0.6rem 1.5rem;
  font-size: 0.85rem; font-weight: 600; cursor: pointer; margin-bottom: 0.75rem;
}

.cancel-link {
  background: none; border: none; color: #54668f;
  font-size: 0.82rem; font-weight: 500; cursor: pointer;
  display: flex; align-items: center; gap: 5px; transition: color 200ms; margin-top: 0.75rem;
}
.cancel-link:hover { color: #1565c0; }

/* Report body */
.report-body { max-width: 900px; margin: 0 auto; padding: 1.25rem; }

/* Cover */
.report-cover {
  background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 45%, #0d47a1 100%);
  border-radius: 14px; padding: 2rem; display: flex;
  align-items: center; justify-content: space-between; gap: 1.5rem;
  margin-bottom: 0; border-bottom-left-radius: 0; border-bottom-right-radius: 0;
  box-shadow: 0 4px 20px rgba(0,0,0,0.12);
}
.report-cover-label { font-size: 0.6rem; text-transform: uppercase; letter-spacing: 0.12em; color: rgba(255,255,255,0.65); margin-bottom: 0.4rem; }
.report-cover-title { font-size: 1.3rem; font-weight: 800; color: white; margin-bottom: 0.6rem; line-height: 1.25; }
.report-cover-meta { display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center; font-size: 0.72rem; color: rgba(255,255,255,0.75); }
.meta-sep { opacity: 0.4; }

.report-cover-score { text-align: center; flex-shrink: 0; }
.cover-grade {
  width: 58px; height: 58px; border-radius: 50%; margin: 0 auto 0.3rem;
  display: flex; align-items: center; justify-content: center;
  font-size: 1.7rem; font-weight: 800; color: white;
  box-shadow: 0 4px 16px rgba(0,0,0,0.3);
}
.grade-A { background: #2e7d32; } .grade-B { background: #558b2f; }
.grade-C { background: #e65100; } .grade-D { background: #c62828; } .grade-E { background: #4a148c; }
.cover-score-val { font-size: 1.5rem; font-weight: 800; color: white; }
.cover-score-val span { font-size: 0.75rem; opacity: 0.65; }
.cover-score-lbl { font-size: 0.6rem; color: rgba(255,255,255,0.6); letter-spacing: 0.08em; }

/* Score strip */
.score-strip {
  display: flex; background: white;
  border-left: 1px solid rgba(0,0,0,0.07); border-right: 1px solid rgba(0,0,0,0.07);
  border-bottom: 1px solid rgba(0,0,0,0.07);
  border-radius: 0 0 12px 12px;
  overflow: hidden; margin-bottom: 1.25rem;
  box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.score-strip-item {
  flex: 1; padding: 0.75rem 1rem;
  border-right: 1px solid rgba(0,0,0,0.06);
}
.score-strip-item:last-child { border-right: none; }
.score-strip-val { font-size: 1.6rem; font-weight: 800; line-height: 1; }
.score-strip-label { font-size: 0.65rem; color: #54668f; margin-top: 2px; margin-bottom: 0.4rem; }
.score-strip-bar { height: 4px; background: rgba(0,0,0,0.06); border-radius: 2px; overflow: hidden; }
.score-strip-fill { height: 100%; border-radius: 2px; transition: width 600ms ease; }

/* Sections */
.report-sections { display: flex; flex-direction: column; gap: 1rem; }

.report-section {
  background: white; border-radius: 12px;
  border: 1px solid rgba(0,0,0,0.07);
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,0.04);
  padding: 1.25rem 1.5rem;
}
.report-section--last {
  background: linear-gradient(135deg, rgba(21,101,192,0.03), rgba(46,125,50,0.03));
}

.section-eyebrow { font-size: 0.6rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.12em; color: #9e9e9e; margin-bottom: 0.2rem; }
.section-title { font-size: 1rem; font-weight: 800; color: #0d1f3c; margin-bottom: 1rem; padding-bottom: 0.6rem; border-bottom: 2px solid rgba(0,0,0,0.06); }
.report-para { font-size: 0.83rem; color: #2a3a5e; line-height: 1.8; margin-bottom: 0.5rem; }
.report-para--highlight { background: rgba(21,101,192,0.04); border-left: 3px solid #1565c0; padding: 0.75rem 1rem; border-radius: 0 8px 8px 0; }

/* Findings */
.finding-row { display: flex; gap: 0.75rem; align-items: flex-start; margin-bottom: 0.65rem; padding: 0.5rem 0.75rem; background: #f8faff; border-radius: 8px; border-left: 3px solid #1565c0; }
.finding-num { font-size: 0.7rem; font-weight: 800; color: #1565c0; min-width: 24px; margin-top: 2px; }
.finding-text { font-size: 0.81rem; color: #2a3a5e; line-height: 1.6; margin: 0; }

/* Risk cards */
.risk-card { border: 1px solid rgba(0,0,0,0.08); border-radius: 10px; overflow: hidden; margin-bottom: 0.85rem; }
.risk-card-header { display: flex; align-items: center; gap: 0.75rem; padding: 0.65rem 1rem; background: rgba(198,40,40,0.04); border-bottom: 1px solid rgba(198,40,40,0.1); }
.risk-rank { width: 28px; height: 28px; border-radius: 50%; background: #c62828; color: white; display: flex; align-items: center; justify-content: center; font-size: 0.75rem; font-weight: 800; flex-shrink: 0; }
.risk-title { font-size: 0.85rem; font-weight: 700; color: #1a0d0d; }
.risk-desc { padding: 0.6rem 1rem; font-size: 0.8rem; color: #3a2020; line-height: 1.65; }
.risk-mitigation { padding: 0.6rem 1rem; background: rgba(46,125,50,0.04); border-top: 1px solid rgba(46,125,50,0.12); }
.risk-mit-label { display: flex; align-items: center; gap: 0.35rem; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #2e7d32; margin-bottom: 0.3rem; }
.risk-mit-text { font-size: 0.78rem; color: #1b5e20; line-height: 1.6; margin: 0; }

/* Action table */
.action-table-wrap { overflow-x: auto; }
.action-table { width: 100%; border-collapse: collapse; font-size: 0.78rem; }
.action-table th { background: #f0f4ff; padding: 0.45rem 0.65rem; text-align: left; font-size: 0.68rem; font-weight: 700; color: #3a4a6b; border-bottom: 2px solid #dde6f7; white-space: nowrap; }
.action-table td { padding: 0.5rem 0.65rem; border-bottom: 1px solid rgba(0,0,0,0.05); vertical-align: middle; }
.action-text { font-size: 0.78rem; color: #2a3a5e; line-height: 1.5; }
.timeline-text { font-size: 0.72rem; color: #54668f; white-space: nowrap; }

.priority-badge { padding: 2px 8px; border-radius: 5px; font-size: 0.65rem; font-weight: 700; color: white; }
.priority--high { background: #c62828; }
.priority--medium { background: #e65100; }
.priority--low { background: #2e7d32; }

.area-badge { padding: 2px 8px; border-radius: 5px; font-size: 0.65rem; font-weight: 600; }
.area--csrd { background: rgba(21,101,192,0.1); color: #1565c0; }
.area--eu_taxonomy { background: rgba(46,125,50,0.1); color: #2e7d32; }
.area--physical_risk { background: rgba(230,81,0,0.1); color: #e65100; }
.area--sfdr { background: rgba(106,27,154,0.1); color: #6a1b9a; }
.area--governance { background: rgba(0,137,123,0.1); color: #00897b; }

/* Footer */
.report-footer { display: flex; align-items: center; justify-content: center; gap: 0.35rem; padding: 1.5rem; font-size: 0.7rem; color: #9e9e9e; }

@media (max-width: 600px) {
  .report-cover { flex-direction: column; align-items: flex-start; }
  .score-strip { flex-wrap: wrap; }
  .score-strip-item { min-width: 50%; }
}
</style>
