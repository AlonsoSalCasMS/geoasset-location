<template>
  <div class="esg-page">

    <!-- ── LOADING ─────────────────────────────────────────────── -->
    <section v-if="store.esgLoading" class="esg-fullscreen">
      <div class="hero-noise hero-noise--green" />
      <div class="esg-loading-content">
        <div class="orbital-wrapper">
          <svg class="orbital-svg" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
            <circle cx="100" cy="100" r="90" fill="none" stroke="rgba(46,125,50,0.08)" stroke-width="1" />
            <circle cx="100" cy="100" r="65" fill="none" stroke="rgba(46,125,50,0.10)" stroke-width="1" />
            <circle cx="100" cy="100" r="40" fill="none" stroke="rgba(46,125,50,0.12)" stroke-width="1" />
            <circle cx="100" cy="100" r="90" fill="none" stroke="url(#esg1)" stroke-width="2.5"
              stroke-dasharray="80 486" stroke-linecap="round" class="orbit-arc orbit-arc--1" />
            <circle cx="100" cy="100" r="65" fill="none" stroke="url(#esg2)" stroke-width="2"
              stroke-dasharray="55 353" stroke-linecap="round" class="orbit-arc orbit-arc--2" />
            <circle cx="100" cy="100" r="40" fill="none" stroke="url(#esg3)" stroke-width="2"
              stroke-dasharray="35 216" stroke-linecap="round" class="orbit-arc orbit-arc--3" />
            <circle cx="100" cy="100" r="12" class="core-pulse-green" />
            <circle cx="100" cy="100" r="6" fill="#2e7d32" />
            <circle r="3.5" fill="#2e7d32" class="orbit-dot">
              <animateMotion dur="6s" repeatCount="indefinite" path="M100,10 A90,90 0 1,1 99.99,10" />
            </circle>
            <circle r="2.5" fill="#00897b" class="orbit-dot">
              <animateMotion dur="4.5s" repeatCount="indefinite" path="M100,35 A65,65 0 1,1 99.99,35" />
            </circle>
            <circle r="2" fill="#1565c0" class="orbit-dot">
              <animateMotion dur="3s" repeatCount="indefinite" path="M100,60 A40,40 0 1,1 99.99,60" />
            </circle>
            <defs>
              <linearGradient id="esg1" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stop-color="#2e7d32" stop-opacity="0" />
                <stop offset="100%" stop-color="#2e7d32" stop-opacity="0.7" />
              </linearGradient>
              <linearGradient id="esg2" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stop-color="#00897b" stop-opacity="0" />
                <stop offset="100%" stop-color="#00897b" stop-opacity="0.6" />
              </linearGradient>
              <linearGradient id="esg3" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stop-color="#1565c0" stop-opacity="0" />
                <stop offset="100%" stop-color="#1565c0" stop-opacity="0.5" />
              </linearGradient>
            </defs>
          </svg>
        </div>

        <h2 class="loading-title">
          Análisis ESG
          <span class="loading-company">{{ store.selectedCompany?.name || store.metadata?.company?.name }}</span>
        </h2>
        <p class="loading-subtitle">Evaluando {{ store.assets.length }} activos con inteligencia artificial</p>
        <div class="loading-pct">{{ loadingPct }}%</div>

        <div class="timeline">
          <div v-for="(step, i) in loadingSteps" :key="i" class="timeline-step"
            :class="{ 'timeline-step--complete': i < loadingStepIdx, 'timeline-step--running': i === loadingStepIdx, 'timeline-step--pending': i > loadingStepIdx }">
            <div v-if="i > 0" class="timeline-connector">
              <div class="timeline-connector-fill" :class="{ filled: i <= loadingStepIdx }" />
            </div>
            <div class="timeline-node">
              <svg v-if="i < loadingStepIdx" class="node-icon" viewBox="0 0 24 24">
                <circle cx="12" cy="12" r="11" fill="#14b86a" />
                <path d="M8 12.5l2.5 2.5 5.5-5.5" fill="none" stroke="white" stroke-width="2"
                  stroke-linecap="round" stroke-linejoin="round" class="check-path" />
              </svg>
              <svg v-else-if="i === loadingStepIdx" class="node-icon" viewBox="0 0 24 24">
                <circle cx="12" cy="12" r="10" fill="none" stroke="rgba(46,125,50,0.2)" stroke-width="2" />
                <circle cx="12" cy="12" r="10" fill="none" stroke="#2e7d32" stroke-width="2"
                  stroke-dasharray="20 43" stroke-linecap="round" class="spinner-arc" />
              </svg>
              <div v-else class="node-pending"><span>{{ i + 1 }}</span></div>
            </div>
            <div class="timeline-label">
              <span class="timeline-name">{{ step }}</span>
            </div>
          </div>
        </div>

        <button class="cancel-link" @click="store.setView('results')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
          Volver a resultados
        </button>
      </div>
    </section>

    <!-- ── ERROR ───────────────────────────────────────────────── -->
    <section v-else-if="store.esgError" class="esg-fullscreen">
      <div class="hero-noise hero-noise--green" />
      <div class="esg-loading-content">
        <svg width="64" height="64" viewBox="0 0 64 64" class="mb-6">
          <circle cx="32" cy="32" r="30" fill="rgba(239,68,68,0.1)" stroke="#ef4444" stroke-width="2" />
          <path d="M32 20v16M32 44v2" stroke="#ef4444" stroke-width="3" stroke-linecap="round" />
        </svg>
        <h2 class="loading-title" style="color:#b91c1c">Error en el análisis</h2>
        <p class="loading-subtitle">{{ store.esgError }}</p>
        <button class="esg-start-btn" @click="runAnalysis" style="margin-top:1.5rem">Reintentar</button>
        <button class="cancel-link" @click="store.setView('results')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
          Volver
        </button>
      </div>
    </section>

    <!-- ── TRIGGER ─────────────────────────────────────────────── -->
    <section v-else-if="!store.esgAnalysis" class="esg-fullscreen">
      <div class="hero-noise hero-noise--green" />
      <div class="esg-loading-content">
        <div class="trigger-icon-wrap">
          <svg width="72" height="72" viewBox="0 0 72 72">
            <circle cx="36" cy="36" r="34" fill="rgba(46,125,50,0.07)" stroke="rgba(46,125,50,0.18)" stroke-width="1.5" />
            <circle cx="36" cy="36" r="22" fill="rgba(46,125,50,0.08)" />
            <path d="M36 20 C28 20 22 26 22 34 C22 44 36 54 36 54 C36 54 50 44 50 34 C50 26 44 20 36 20Z"
              fill="rgba(46,125,50,0.15)" stroke="#2e7d32" stroke-width="1.5" />
            <circle cx="36" cy="34" r="5" fill="#2e7d32" />
          </svg>
        </div>

        <h2 class="loading-title">Análisis ESG de Activos</h2>
        <p class="loading-subtitle" style="max-width:480px;text-align:center">
          Evalúa el riesgo físico climático, la exposición regulatoria y el perfil de sostenibilidad
          de los <strong>{{ store.assets.length }} activos</strong> identificados.
        </p>

        <div class="trigger-chips">
          <span class="trigger-chip chip--green">CSRD / ESRS E1·E3·E4</span>
          <span class="trigger-chip chip--blue">Taxonomía UE · DNSH</span>
          <span class="trigger-chip chip--orange">Riesgo Físico Climático</span>
          <span class="trigger-chip chip--purple">SFDR PAI 7·8·14</span>
          <span class="trigger-chip chip--teal">EBA GL/2020/06</span>
        </div>

        <button class="esg-start-btn" @click="runAnalysis">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" style="flex-shrink:0">
            <path d="M8 5v14l11-7z" />
          </svg>
          Iniciar análisis ESG
        </button>

        <button class="cancel-link" @click="store.setView('results')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M19 12H5M12 19l-7-7 7-7" />
          </svg>
          Volver a resultados
        </button>
      </div>
    </section>

    <!-- ── DASHBOARD ───────────────────────────────────────────── -->
    <div v-else class="dashboard">

      <!-- Top bar -->
      <div class="dash-topbar">
        <div class="dash-topbar-left">
          <button class="back-btn" @click="store.setView('results')">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
              <path d="M19 12H5M12 19l-7-7 7-7" />
            </svg>
          </button>
          <div>
            <div class="dash-overline">Análisis ESG</div>
            <div class="dash-company">{{ store.selectedCompany?.name || store.metadata?.company?.name }}</div>
          </div>
        </div>
        <div class="dash-topbar-right">
          <span class="dash-chip">{{ store.assets.length }} activos</span>
          <button class="report-btn" @click="goToReport">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="16" y1="13" x2="8" y2="13" />
              <line x1="16" y1="17" x2="8" y2="17" />
            </svg>
            Generar Informe
          </button>
        </div>
      </div>

      <!-- KPI row -->
      <div class="kpi-row">
        <!-- Overall score -->
        <div class="kpi-overall">
          <div class="kpi-grade" :class="`grade-${store.esgAnalysis.grade}`">{{ store.esgAnalysis.grade }}</div>
          <div>
            <div class="kpi-score">{{ store.esgAnalysis.overall_score }}<span class="kpi-score-unit">/100</span></div>
            <div class="kpi-score-label">Puntuación ESG Global</div>
          </div>
        </div>
        <!-- 4 dimensions -->
        <div v-for="dim in dimensions" :key="dim.key" class="kpi-dim">
          <div class="kpi-dim-top">
            <div class="kpi-dim-icon" :style="{ background: dim.bg }">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round">
                <path :d="dim.svgPath" />
              </svg>
            </div>
            <div class="kpi-dim-score" :style="{ color: dim.color }">{{ dim.score }}</div>
          </div>
          <div class="kpi-dim-label">{{ dim.label }}</div>
          <div class="kpi-dim-weight">Peso {{ (dim.weight * 100).toFixed(0) }}%</div>
          <div class="kpi-bar-track">
            <div class="kpi-bar-fill" :style="{ width: dim.score + '%', background: dim.color }" />
          </div>
        </div>
      </div>

      <!-- Main dashboard grid -->
      <div class="dash-grid">

        <!-- Physical risk table -->
        <div class="dash-card dash-card--wide">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#f57c00"></div>
            Riesgo Físico por Activo
            <span class="dash-card-badge">{{ store.esgAnalysis.asset_risks.length }} activos</span>
          </div>
          <div class="risk-table-wrap">
            <table class="risk-table">
              <thead>
                <tr>
                  <th>Activo</th>
                  <th>Cat.</th>
                  <th>Provincia</th>
                  <th>
                    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
                      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z M12 16v-4M12 8h.01" />
                    </svg>
                    Hídrico
                  </th>
                  <th>Inund.</th>
                  <th>Calor</th>
                  <th>Incend.</th>
                  <th>Score</th>
                  <th>Nivel</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="r in store.esgAnalysis.asset_risks" :key="r.asset_id">
                  <td class="risk-name">{{ r.asset_name }}</td>
                  <td><span class="cat-badge">{{ r.category }}</span></td>
                  <td class="risk-province">{{ r.province }}</td>
                  <td><risk-pill :level="r.water_stress" /></td>
                  <td><risk-pill :level="r.flood_risk" /></td>
                  <td><risk-pill :level="r.heat_risk" /></td>
                  <td><risk-pill :level="r.wildfire_risk" /></td>
                  <td>
                    <span class="risk-pct" :style="{ color: scoreColor(r.composite_risk_score) }">
                      {{ (r.composite_risk_score * 100).toFixed(0) }}%
                    </span>
                  </td>
                  <td><tier-badge :tier="r.risk_tier" /></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- Key insights -->
        <div class="dash-card">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#1565c0"></div>
            Hallazgos Clave
          </div>
          <div class="insights-list">
            <div v-for="(ins, i) in store.esgAnalysis.key_insights" :key="i"
              class="insight-row" :class="`insight--${ins.type}`">
              <div class="insight-icon">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round">
                  <path :d="insightSvg(ins.type)" />
                </svg>
              </div>
              <div>
                <div class="insight-title">{{ ins.title }}</div>
                <div class="insight-detail">{{ ins.detail }}</div>
              </div>
            </div>
          </div>
        </div>

        <!-- CSRD readiness -->
        <div class="dash-card">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#1565c0"></div>
            CSRD / ESRS
            <span class="dash-card-badge" :class="csrdBadgeClass">{{ store.esgAnalysis.csrd_readiness.status }}</span>
          </div>
          <div class="csrd-score-row">
            <div class="csrd-ring" :class="csrdRingClass">{{ store.esgAnalysis.csrd_readiness.score }}</div>
            <div class="csrd-score-label">Preparación</div>
          </div>
          <div v-for="esrs in esrsItems" :key="esrs.key" class="esrs-row">
            <div class="esrs-row-top">
              <span class="esrs-tag" :class="`esrs-tag--${esrs.statusKey}`">{{ esrs.status }}</span>
              <span class="esrs-label">{{ esrs.label }}</span>
            </div>
            <div class="esrs-gap">{{ esrs.gap }}</div>
          </div>
        </div>

        <!-- EU Taxonomy -->
        <div class="dash-card">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#2e7d32"></div>
            Taxonomía UE
          </div>
          <div class="taxonomy-big" :style="{ color: taxonomyColor }">
            {{ store.esgAnalysis.eu_taxonomy.alignment_pct }}%
            <span class="taxonomy-big-label">alineación estimada</span>
          </div>
          <div class="bar-track mt-2 mb-3">
            <div class="bar-fill" :style="{ width: store.esgAnalysis.eu_taxonomy.alignment_pct + '%', background: taxonomyColor }" />
          </div>
          <div v-if="store.esgAnalysis.eu_taxonomy.aligned_categories.length" class="tax-cats">
            <span v-for="cat in store.esgAnalysis.eu_taxonomy.aligned_categories" :key="cat" class="cat-badge cat-badge--green">{{ cat }}</span>
          </div>
          <div class="tax-dnsh">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#e65100" stroke-width="2.5" stroke-linecap="round">
              <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z M12 9v4M12 17h.01" />
            </svg>
            {{ store.esgAnalysis.eu_taxonomy.dnsh_bottleneck }}
          </div>
        </div>

        <!-- Province heatmap -->
        <div class="dash-card">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#c62828"></div>
            Mapa de Riesgo por Provincia
          </div>
          <div v-for="prov in sortedProvinces" :key="prov.province" class="prov-row">
            <div class="prov-row-top">
              <span class="prov-name">{{ prov.province }}</span>
              <span class="prov-assets">{{ prov.asset_count }} act.</span>
              <span class="prov-pct" :style="{ color: scoreColor(prov.avg_risk) }">{{ (prov.avg_risk * 100).toFixed(0) }}%</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill" :style="{ width: (prov.avg_risk * 100) + '%', background: scoreColor(prov.avg_risk) }" />
            </div>
            <div class="prov-risk-label">{{ prov.dominant_risk }}</div>
          </div>
        </div>

        <!-- Concentration + Carbon + SFDR -->
        <div class="dash-card">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#6a1b9a"></div>
            Inteligencia de Riesgos
          </div>

          <!-- Concentration -->
          <div class="metric-block">
            <div class="metric-row">
              <div class="metric-big">{{ (store.esgAnalysis.concentration_risk.hhi_index * 100).toFixed(0) }}<span class="metric-unit">HHI</span></div>
              <div>
                <div class="metric-badge" :class="concentrationBadgeClass">{{ store.esgAnalysis.concentration_risk.concentration_level }}</div>
                <div class="metric-sub">Concentración en <strong>{{ store.esgAnalysis.concentration_risk.top_province }}</strong></div>
              </div>
            </div>
          </div>

          <div class="metric-divider" />

          <!-- Carbon -->
          <div class="metric-block">
            <div class="metric-row">
              <div class="metric-big" :style="{ color: emissionColor }">{{ store.esgAnalysis.carbon_exposure.high_emission_asset_pct }}%<span class="metric-unit">CO₂</span></div>
              <div>
                <div class="metric-badge" :class="transitionBadgeClass">Transición {{ store.esgAnalysis.carbon_exposure.transition_risk_level }}</div>
                <div class="metric-sub">activos de alta emisión</div>
              </div>
            </div>
          </div>

          <div class="metric-divider" />

          <!-- SFDR PAI -->
          <div class="sfdr-label">SFDR – Principal Adverse Impacts</div>
          <div v-for="pai in paiItems" :key="pai.label" class="pai-row">
            <div class="pai-row-top">
              <span class="pai-label">{{ pai.label }}</span>
              <span class="pai-val" :style="{ color: scoreColor(pai.value) }">{{ (pai.value * 100).toFixed(0) }}</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill" :style="{ width: (pai.value * 100) + '%', background: scoreColor(pai.value) }" />
            </div>
          </div>
          <div class="pai-summary">{{ store.esgAnalysis.sfdr_pai.summary }}</div>
        </div>

        <!-- Executive summary -->
        <div class="dash-card dash-card--wide">
          <div class="dash-card-header">
            <div class="dash-card-dot" style="background:#0d47a1"></div>
            Resumen Ejecutivo
            <button class="report-btn report-btn--inline" @click="goToReport">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                <polyline points="14 2 14 8 20 8" />
              </svg>
              Ver informe completo
            </button>
          </div>
          <p class="exec-summary">{{ store.esgAnalysis.executive_summary }}</p>
        </div>

      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, defineComponent, h } from 'vue';
import { useAppStore } from '@/stores/store';
import { analyzeEsg } from '@/services/backend';

// ── Inline sub-components ───────────────────────────────────────
const RISK_COLORS: Record<string, string> = { LOW: '#2e7d32', MEDIUM: '#e65100', HIGH: '#c62828', VERY_HIGH: '#4a148c' };

const RiskPill = defineComponent({
  props: ['level'],
  setup(props) {
    return () => h('span', {
      style: {
        display: 'inline-block', padding: '1px 6px', borderRadius: '4px',
        fontSize: '0.68rem', fontWeight: '700', color: 'white',
        background: RISK_COLORS[props.level] || '#9e9e9e',
        letterSpacing: '0.02em',
      }
    }, props.level === 'VERY_HIGH' ? 'V.HIGH' : props.level);
  },
});

const TIER_COLORS: Record<string, string> = { LOW: '#2e7d32', MEDIUM: '#e65100', HIGH: '#c62828', CRITICAL: '#4a148c' };
const TierBadge = defineComponent({
  props: ['tier'],
  setup(props) {
    return () => h('span', {
      style: {
        display: 'inline-block', padding: '2px 8px', borderRadius: '6px',
        fontSize: '0.68rem', fontWeight: '700', color: 'white',
        background: TIER_COLORS[props.tier] || '#9e9e9e',
      }
    }, props.tier);
  },
});

// ── Store + state ───────────────────────────────────────────────
const store = useAppStore();

const loadingSteps = [
  'Evaluando riesgos físicos',
  'Analizando exposición CSRD',
  'Calculando Taxonomía UE',
  'Estimando SFDR PAI',
  'Generando score ESG',
];
const loadingStepIdx = ref(0);
const loadingPct = computed(() => Math.round((loadingStepIdx.value / (loadingSteps.length - 1)) * 100));

let stepInterval: ReturnType<typeof setInterval> | null = null;

const runAnalysis = async () => {
  store.setEsgLoading(true);
  store.setEsgError(null);
  loadingStepIdx.value = 0;
  stepInterval = setInterval(() => {
    if (loadingStepIdx.value < loadingSteps.length - 1) loadingStepIdx.value++;
  }, 12000);
  try {
    const companyName = store.selectedCompany?.name || store.metadata?.company?.name || '';
    const result = await analyzeEsg(companyName, store.assets);
    if (result && (result as any).error) throw new Error((result as any).error);
    store.setEsgAnalysis(result);
  } catch (e: any) {
    store.setEsgError(e?.message || 'Error desconocido');
  } finally {
    if (stepInterval) clearInterval(stepInterval);
    store.setEsgLoading(false);
  }
};

const goToReport = () => store.setView('report');

// ── Dimensions ──────────────────────────────────────────────────
const dimensions = computed(() => {
  if (!store.esgAnalysis) return [];
  const s = store.esgAnalysis.scores;
  return [
    { key: 'pr', label: 'Riesgo Físico', score: s.physical_risk.score, weight: s.physical_risk.weight, color: '#c62828', bg: 'rgba(198,40,40,0.1)', svgPath: 'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z' },
    { key: 'ei', label: 'Imp. Ambiental', score: s.environmental_impact.score, weight: s.environmental_impact.weight, color: '#2e7d32', bg: 'rgba(46,125,50,0.1)', svgPath: 'M17 8C8 10 5.9 16.17 3.82 19.5c1.1.5 2.28.67 3.28.5C8.5 19 9.5 17 12 17s3.5 2 5.5 2 3.5-2 3.5-2V9.5C18.5 8 17 8 17 8z' },
    { key: 'rt', label: 'Trans. Regulatoria', score: s.regulatory_transition.score, weight: s.regulatory_transition.weight, color: '#1565c0', bg: 'rgba(21,101,192,0.1)', svgPath: 'M12 1v22M5 6l7-5 7 5M5 18l7 5 7-5M5 6v12M19 6v12' },
    { key: 're', label: 'Resiliencia', score: s.resilience.score, weight: s.resilience.weight, color: '#6a1b9a', bg: 'rgba(106,27,154,0.1)', svgPath: 'M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z' },
  ];
});

// ── Helpers ─────────────────────────────────────────────────────
const scoreColor = (val: number) => {
  if (val >= 0.65) return '#c62828';
  if (val >= 0.35) return '#e65100';
  return '#2e7d32';
};

const esrsItems = computed(() => {
  if (!store.esgAnalysis) return [];
  const r = store.esgAnalysis.csrd_readiness;
  return [
    { key: 'e1', label: 'E1 Clima', status: r.esrs_e1.status, statusKey: r.esrs_e1.status.toLowerCase().replace('_', '-'), gap: r.esrs_e1.key_gap },
    { key: 'e3', label: 'E3 Agua', status: r.esrs_e3.status, statusKey: r.esrs_e3.status.toLowerCase().replace('_', '-'), gap: r.esrs_e3.key_gap },
    { key: 'e4', label: 'E4 Biodiversidad', status: r.esrs_e4.status, statusKey: r.esrs_e4.status.toLowerCase().replace('_', '-'), gap: r.esrs_e4.key_gap },
  ];
});

const csrdBadgeClass = computed(() => {
  const s = store.esgAnalysis?.csrd_readiness.status;
  if (s === 'COMPLIANT' || s === 'ADVANCED') return 'badge--green';
  if (s === 'PARTIAL') return 'badge--orange';
  return 'badge--red';
});

const csrdRingClass = computed(() => {
  const s = store.esgAnalysis?.csrd_readiness.status;
  if (s === 'COMPLIANT') return 'ring--green';
  if (s === 'ADVANCED') return 'ring--lime';
  if (s === 'PARTIAL') return 'ring--orange';
  return 'ring--red';
});

const taxonomyColor = computed(() => {
  const p = store.esgAnalysis?.eu_taxonomy.alignment_pct ?? 0;
  return p >= 50 ? '#2e7d32' : p >= 20 ? '#e65100' : '#c62828';
});

const sortedProvinces = computed(() =>
  [...(store.esgAnalysis?.province_summary || [])].sort((a, b) => b.avg_risk - a.avg_risk).slice(0, 10)
);

const concentrationBadgeClass = computed(() => {
  const l = store.esgAnalysis?.concentration_risk.concentration_level;
  if (l === 'VERY_HIGH' || l === 'HIGH') return 'badge--red';
  if (l === 'MODERATE') return 'badge--orange';
  return 'badge--green';
});

const emissionColor = computed(() => {
  const p = store.esgAnalysis?.carbon_exposure.high_emission_asset_pct ?? 0;
  return p >= 50 ? '#c62828' : p >= 25 ? '#e65100' : '#2e7d32';
});

const transitionBadgeClass = computed(() => {
  const l = store.esgAnalysis?.carbon_exposure.transition_risk_level;
  return l === 'HIGH' ? 'badge--red' : l === 'MEDIUM' ? 'badge--orange' : 'badge--green';
});

const paiItems = computed(() => {
  if (!store.esgAnalysis) return [];
  const p = store.esgAnalysis.sfdr_pai;
  return [
    { label: 'PAI 7 GEI Scope 1+2', value: p.pai_7_score },
    { label: 'PAI 8 Huella Carbono', value: p.pai_8_score },
    { label: 'PAI 14 Biodiversidad', value: p.pai_14_score },
  ];
});

const insightSvg = (type: string) => {
  const paths: Record<string, string> = {
    risk: 'M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z M12 9v4M12 17h.01',
    opportunity: 'M23 6l-9.5 9.5-5-5L1 18 M17 6h6v6',
    regulatory: 'M12 1v22M5 6l7-5 7 5M5 18l7 5 7-5',
    info: 'M12 16v-4M12 8h.01 M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z',
  };
  return paths[type] || paths.info;
};
</script>

<style scoped>
/* ── Page ── */
.esg-page {
  min-height: 100vh;
  background: #f4f7fb;
  overflow-y: auto;
}

/* ── Fullscreen states (loading / trigger / error) ── */
.esg-fullscreen {
  position: relative;
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}

.hero-noise--green {
  position: absolute;
  inset: 0;
  background:
    radial-gradient(circle at 18% 18%, rgba(46, 125, 50, 0.14), transparent 36%),
    radial-gradient(circle at 82% 10%, rgba(0, 137, 123, 0.12), transparent 36%),
    radial-gradient(circle at 50% 88%, rgba(21, 101, 192, 0.08), transparent 35%),
    linear-gradient(160deg, rgba(255,255,255,0.9) 0%, rgba(240,248,240,0.95) 50%, rgba(240,245,255,0.9) 100%);
  pointer-events: none;
}

.esg-loading-content {
  position: relative;
  z-index: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  width: min(640px, 100%);
  padding: 0 1.5rem;
}

/* Orbital */
.orbital-wrapper {
  width: 170px;
  height: 170px;
  margin-bottom: 1.5rem;
  animation: fadeScaleIn 600ms ease-out both;
}
.orbital-svg { width: 100%; height: 100%; }
.orbit-arc { transform-origin: center; }
.orbit-arc--1 { animation: orbitSpin 8s linear infinite; }
.orbit-arc--2 { animation: orbitSpin 6s linear infinite reverse; }
.orbit-arc--3 { animation: orbitSpin 4s linear infinite; }
.core-pulse-green { fill: rgba(46,125,50,0.12); animation: pulseGreen 2s ease-in-out infinite; }
@keyframes orbitSpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
@keyframes pulseGreen { 0%,100% { r: 12; opacity: 0.4; } 50% { r: 18; opacity: 0.08; } }
@keyframes fadeScaleIn { from { opacity: 0; transform: scale(0.85); } to { opacity: 1; transform: scale(1); } }

/* Loading typography */
.loading-title {
  font-size: clamp(1.4rem, 3vw, 2rem);
  font-weight: 700;
  color: #0e1a37;
  margin-bottom: 0.3rem;
  animation: riseIn 500ms ease-out both;
}
.loading-company {
  display: block;
  background: linear-gradient(90deg, #2e7d32, #00897b);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}
.loading-subtitle { color: #54668f; font-size: 0.95rem; margin-bottom: 0.5rem; animation: riseIn 600ms ease-out both; }
.loading-pct { font-size: 2.4rem; font-weight: 800; color: #2e7d32; letter-spacing: -0.03em; margin-bottom: 1.8rem; font-variant-numeric: tabular-nums; animation: riseIn 650ms ease-out both; }
@keyframes riseIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

/* Timeline */
.timeline {
  display: flex;
  align-items: flex-start;
  gap: 0;
  width: 100%;
  max-width: 520px;
  animation: riseIn 750ms ease-out both;
}
.timeline-step { flex: 1; display: flex; flex-direction: column; align-items: center; position: relative; }
.timeline-connector { position: absolute; top: 14px; right: 50%; width: 100%; height: 2px; background: rgba(46,125,50,0.12); z-index: 0; }
.timeline-connector-fill { height: 100%; width: 0%; background: #14b86a; border-radius: 1px; transition: width 600ms ease; }
.timeline-connector-fill.filled { width: 100%; }
.timeline-node { width: 28px; height: 28px; position: relative; z-index: 1; display: flex; align-items: center; justify-content: center; margin-bottom: 0.45rem; }
.node-icon { width: 28px; height: 28px; }
.spinner-arc { transform-origin: center; animation: orbitSpin 1s linear infinite; }
.check-path { stroke-dasharray: 20; stroke-dashoffset: 20; animation: drawCheck 400ms ease-out forwards; }
@keyframes drawCheck { to { stroke-dashoffset: 0; } }
.node-pending { width: 28px; height: 28px; border-radius: 50%; border: 2px solid rgba(46,125,50,0.2); background: rgba(255,255,255,0.9); display: flex; align-items: center; justify-content: center; }
.node-pending span { font-size: 0.68rem; font-weight: 600; color: #8fa0bf; }
.timeline-label { display: flex; flex-direction: column; align-items: center; }
.timeline-name { font-size: 0.68rem; font-weight: 500; color: #8fa0bf; transition: color 300ms; text-align: center; max-width: 80px; line-height: 1.3; }
.timeline-step--complete .timeline-name { color: #12254d; }
.timeline-step--running .timeline-name { color: #2e7d32; font-weight: 600; }

/* Trigger screen */
.trigger-icon-wrap { margin-bottom: 1.5rem; }
.trigger-chips { display: flex; flex-wrap: wrap; justify-content: center; gap: 0.5rem; margin: 1rem 0 2rem; }
.trigger-chip { padding: 4px 12px; border-radius: 100px; font-size: 0.75rem; font-weight: 600; }
.chip--green { background: rgba(46,125,50,0.1); color: #2e7d32; }
.chip--blue { background: rgba(21,101,192,0.1); color: #1565c0; }
.chip--orange { background: rgba(230,81,0,0.1); color: #e65100; }
.chip--purple { background: rgba(106,27,154,0.1); color: #6a1b9a; }
.chip--teal { background: rgba(0,137,123,0.1); color: #00897b; }

.esg-start-btn {
  display: inline-flex; align-items: center; gap: 0.5rem;
  background: linear-gradient(135deg, #2e7d32, #1b5e20);
  color: white; border: none; border-radius: 12px;
  padding: 0.75rem 2rem; font-size: 0.95rem; font-weight: 600;
  cursor: pointer; box-shadow: 0 4px 14px rgba(46,125,50,0.35);
  transition: transform 150ms, box-shadow 150ms;
}
.esg-start-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(46,125,50,0.4); }

.cancel-link {
  margin-top: 1.8rem; background: none; border: none; color: #54668f;
  font-size: 0.85rem; font-weight: 500; cursor: pointer;
  display: flex; align-items: center; gap: 6px; transition: color 200ms;
}
.cancel-link:hover { color: #2e7d32; }

/* ── Dashboard ── */
.dashboard { min-height: 100vh; display: flex; flex-direction: column; }

.dash-topbar {
  background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 45%, #0d47a1 100%);
  padding: 0.85rem 1.5rem;
  display: flex; align-items: center; justify-content: space-between;
}
.dash-topbar-left { display: flex; align-items: center; gap: 0.75rem; }
.dash-topbar-right { display: flex; align-items: center; gap: 0.75rem; }

.back-btn {
  width: 32px; height: 32px; border-radius: 8px;
  background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2);
  color: white; display: flex; align-items: center; justify-content: center;
  cursor: pointer; transition: background 150ms;
}
.back-btn:hover { background: rgba(255,255,255,0.25); }

.dash-overline { font-size: 0.62rem; text-transform: uppercase; letter-spacing: 0.1em; color: rgba(255,255,255,0.65); }
.dash-company { font-size: 1.05rem; font-weight: 700; color: white; }
.dash-chip { padding: 3px 10px; border-radius: 100px; border: 1px solid rgba(255,255,255,0.3); color: rgba(255,255,255,0.85); font-size: 0.72rem; }

.report-btn {
  display: inline-flex; align-items: center; gap: 0.4rem;
  background: rgba(255,255,255,0.9); color: #0d1f3c;
  border: none; border-radius: 8px; padding: 0.4rem 0.9rem;
  font-size: 0.78rem; font-weight: 600; cursor: pointer;
  transition: background 150ms;
}
.report-btn:hover { background: white; }

.report-btn--inline {
  margin-left: auto; background: rgba(13,31,60,0.06); color: #0d47a1;
  padding: 0.3rem 0.75rem; font-size: 0.72rem;
}
.report-btn--inline:hover { background: rgba(13,31,60,0.1); }

/* KPI row */
.kpi-row {
  display: grid;
  grid-template-columns: 160px repeat(4, 1fr);
  gap: 0;
  background: white;
  border-bottom: 1px solid rgba(0,0,0,0.07);
}

.kpi-overall {
  display: flex; align-items: center; gap: 0.75rem;
  padding: 1rem 1.25rem;
  border-right: 1px solid rgba(0,0,0,0.07);
}

.kpi-grade {
  width: 52px; height: 52px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 1.5rem; font-weight: 800; color: white; flex-shrink: 0;
}
.grade-A { background: linear-gradient(135deg, #2e7d32, #1b5e20); }
.grade-B { background: linear-gradient(135deg, #558b2f, #33691e); }
.grade-C { background: linear-gradient(135deg, #e65100, #bf360c); }
.grade-D { background: linear-gradient(135deg, #c62828, #b71c1c); }
.grade-E { background: linear-gradient(135deg, #6a1b9a, #4a148c); }

.kpi-score { font-size: 1.7rem; font-weight: 800; color: #0d1f3c; line-height: 1; }
.kpi-score-unit { font-size: 0.8rem; color: #9e9e9e; font-weight: 400; }
.kpi-score-label { font-size: 0.65rem; color: #9e9e9e; margin-top: 2px; }

.kpi-dim {
  padding: 0.85rem 1rem;
  border-right: 1px solid rgba(0,0,0,0.07);
}
.kpi-dim:last-child { border-right: none; }

.kpi-dim-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.2rem; }
.kpi-dim-icon { width: 26px; height: 26px; border-radius: 6px; display: flex; align-items: center; justify-content: center; }
.kpi-dim-icon svg { width: 14px; height: 14px; }
.kpi-dim-score { font-size: 1.4rem; font-weight: 800; }
.kpi-dim-label { font-size: 0.72rem; font-weight: 600; color: #2a3a5e; }
.kpi-dim-weight { font-size: 0.62rem; color: #9e9e9e; margin-bottom: 0.4rem; }
.kpi-bar-track { height: 4px; background: rgba(0,0,0,0.06); border-radius: 2px; overflow: hidden; }
.kpi-bar-fill { height: 100%; border-radius: 2px; transition: width 800ms ease; }

/* Dashboard grid */
.dash-grid {
  display: grid;
  grid-template-columns: minmax(0, 1.6fr) minmax(0, 1fr);
  gap: 0.75rem;
  padding: 0.75rem;
  flex: 1;
}

.dash-card {
  background: white;
  border-radius: 12px;
  border: 1px solid rgba(0,0,0,0.07);
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}

.dash-card--wide {
  grid-column: span 2;
}

.dash-card-header {
  display: flex; align-items: center; gap: 0.5rem;
  padding: 0.65rem 0.9rem;
  border-bottom: 1px solid rgba(0,0,0,0.05);
  font-size: 0.78rem; font-weight: 700; color: #1a2540;
  background: rgba(0,0,0,0.015);
}

.dash-card-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.dash-card-badge {
  margin-left: auto; padding: 1px 8px; border-radius: 100px;
  font-size: 0.65rem; font-weight: 600; background: rgba(0,0,0,0.07); color: #3a4a6b;
}
.badge--green { background: rgba(46,125,50,0.12); color: #2e7d32 !important; }
.badge--orange { background: rgba(230,81,0,0.12); color: #e65100 !important; }
.badge--red { background: rgba(198,40,40,0.12); color: #c62828 !important; }

/* Risk table */
.risk-table-wrap { overflow-x: auto; max-height: 320px; overflow-y: auto; }
.risk-table { width: 100%; border-collapse: collapse; font-size: 0.75rem; }
.risk-table th {
  position: sticky; top: 0; background: #f8faff;
  padding: 0.45rem 0.6rem; text-align: left;
  font-size: 0.68rem; font-weight: 700; color: #3a4a6b;
  border-bottom: 1px solid rgba(0,0,0,0.07);
  white-space: nowrap;
}
.risk-table th svg { vertical-align: middle; margin-right: 2px; }
.risk-table td { padding: 0.35rem 0.6rem; border-bottom: 1px solid rgba(0,0,0,0.04); white-space: nowrap; }
.risk-table tbody tr:hover { background: rgba(0,0,0,0.02); }
.risk-name { max-width: 180px; overflow: hidden; text-overflow: ellipsis; font-weight: 500; color: #12254d; }
.risk-province { color: #54668f; }
.risk-pct { font-weight: 700; }
.cat-badge {
  display: inline-block; padding: 1px 6px; border-radius: 4px;
  font-size: 0.65rem; font-weight: 700; background: rgba(21,101,192,0.1); color: #1565c0;
}
.cat-badge--green { background: rgba(46,125,50,0.1); color: #2e7d32; }

/* Insights */
.insights-list { padding: 0.5rem 0.75rem; display: flex; flex-direction: column; gap: 0.5rem; }
.insight-row {
  display: flex; align-items: flex-start; gap: 0.6rem;
  padding: 0.55rem 0.6rem; border-radius: 8px; border: 1px solid transparent;
}
.insight--risk { background: rgba(198,40,40,0.04); border-color: rgba(198,40,40,0.12); }
.insight--risk .insight-icon { color: #c62828; }
.insight--opportunity { background: rgba(46,125,50,0.04); border-color: rgba(46,125,50,0.12); }
.insight--opportunity .insight-icon { color: #2e7d32; }
.insight--regulatory { background: rgba(21,101,192,0.04); border-color: rgba(21,101,192,0.12); }
.insight--regulatory .insight-icon { color: #1565c0; }
.insight--info { background: rgba(230,81,0,0.04); border-color: rgba(230,81,0,0.12); }
.insight--info .insight-icon { color: #e65100; }
.insight-icon { flex-shrink: 0; margin-top: 1px; }
.insight-title { font-size: 0.76rem; font-weight: 700; color: #12254d; margin-bottom: 2px; }
.insight-detail { font-size: 0.7rem; color: #54668f; line-height: 1.45; }

/* CSRD */
.csrd-score-row { display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem 0.9rem 0.5rem; }
.csrd-ring {
  width: 50px; height: 50px; border-radius: 50%; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  font-size: 1.1rem; font-weight: 800; color: white;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
.ring--green { background: linear-gradient(135deg, #2e7d32, #1b5e20); }
.ring--lime { background: linear-gradient(135deg, #558b2f, #33691e); }
.ring--orange { background: linear-gradient(135deg, #e65100, #bf360c); }
.ring--red { background: linear-gradient(135deg, #c62828, #b71c1c); }
.csrd-score-label { font-size: 0.72rem; color: #54668f; }
.esrs-row { padding: 0.45rem 0.9rem; border-top: 1px solid rgba(0,0,0,0.04); }
.esrs-row-top { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 2px; }
.esrs-tag {
  padding: 1px 6px; border-radius: 4px;
  font-size: 0.65rem; font-weight: 700; color: white; flex-shrink: 0;
}
.esrs-tag--not_started, .esrs-tag--not-started { background: #c62828; }
.esrs-tag--partial { background: #e65100; }
.esrs-tag--advanced { background: #558b2f; }
.esrs-label { font-size: 0.72rem; font-weight: 600; color: #2a3a5e; }
.esrs-gap { font-size: 0.67rem; color: #54668f; }

/* Taxonomy */
.taxonomy-big {
  padding: 0.5rem 0.9rem 0;
  font-size: 2.2rem; font-weight: 800; line-height: 1;
}
.taxonomy-big-label { display: block; font-size: 0.65rem; color: #9e9e9e; font-weight: 400; }
.tax-cats { display: flex; flex-wrap: wrap; gap: 0.3rem; padding: 0 0.9rem 0.5rem; }
.tax-dnsh {
  display: flex; align-items: flex-start; gap: 0.35rem;
  padding: 0.5rem 0.9rem; margin: 0; font-size: 0.68rem; color: #54668f;
  background: rgba(230,81,0,0.04); border-top: 1px solid rgba(0,0,0,0.05);
}

/* Bars */
.bar-track { height: 5px; background: rgba(0,0,0,0.06); border-radius: 3px; overflow: hidden; }
.bar-fill { height: 100%; border-radius: 3px; transition: width 600ms ease; }
.mt-2 { margin-top: 0.5rem; }
.mb-3 { margin-bottom: 0.75rem; }

/* Province heatmap */
.prov-row { padding: 0.45rem 0.9rem; border-top: 1px solid rgba(0,0,0,0.04); }
.prov-row-top { display: flex; align-items: center; margin-bottom: 3px; }
.prov-name { font-size: 0.75rem; font-weight: 600; color: #12254d; flex: 1; }
.prov-assets { font-size: 0.65rem; color: #9e9e9e; margin-right: 0.5rem; }
.prov-pct { font-size: 0.72rem; font-weight: 700; }
.prov-risk-label { font-size: 0.62rem; color: #9e9e9e; margin-top: 3px; }

/* Risk intel block */
.metric-block { padding: 0.65rem 0.9rem; }
.metric-row { display: flex; align-items: flex-start; gap: 0.75rem; }
.metric-big { font-size: 1.8rem; font-weight: 800; color: #0d1f3c; line-height: 1; }
.metric-unit { font-size: 0.7rem; color: #9e9e9e; font-weight: 400; margin-left: 2px; }
.metric-badge {
  display: inline-block; padding: 2px 8px; border-radius: 5px;
  font-size: 0.65rem; font-weight: 700; margin-bottom: 3px;
}
.metric-sub { font-size: 0.7rem; color: #54668f; }
.metric-divider { height: 1px; background: rgba(0,0,0,0.06); }

.sfdr-label { padding: 0.5rem 0.9rem 0.3rem; font-size: 0.7rem; font-weight: 700; color: #3a4a6b; }
.pai-row { padding: 0.3rem 0.9rem; }
.pai-row-top { display: flex; align-items: center; margin-bottom: 3px; }
.pai-label { font-size: 0.68rem; color: #54668f; flex: 1; }
.pai-val { font-size: 0.72rem; font-weight: 700; }
.pai-summary { padding: 0.4rem 0.9rem 0.65rem; font-size: 0.68rem; color: #54668f; line-height: 1.5; }

/* Exec summary */
.exec-summary { padding: 0.75rem 1rem; font-size: 0.82rem; color: #2a3a5e; line-height: 1.75; }

@media (max-width: 960px) {
  .kpi-row { grid-template-columns: repeat(3, 1fr); }
  .kpi-overall { grid-column: span 3; border-right: none; border-bottom: 1px solid rgba(0,0,0,0.07); }
  .dash-grid { grid-template-columns: 1fr; }
  .dash-card--wide { grid-column: span 1; }
}

@media (max-width: 600px) {
  .timeline { flex-direction: column; align-items: stretch; max-width: 280px; }
  .timeline-step { flex-direction: row; align-items: center; gap: 0.75rem; }
  .timeline-connector { position: static; width: 2px; height: 20px; margin-left: 13px; }
  .timeline-label { align-items: flex-start; }
  .kpi-row { grid-template-columns: 1fr 1fr; }
  .kpi-overall { grid-column: span 2; }
}
</style>
