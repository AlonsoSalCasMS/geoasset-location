<template>
  <div style="width: 100%; height: 100%; position: relative">
    <div ref="mapContainer" class="map-container" style="width: 100%; height: 100%"></div>

    <!-- Legend + color-mode toggle overlay -->
    <div class="legend-overlay">
      <v-btn-toggle
        v-model="colorMode"
        density="compact"
        variant="outlined"
        color="primary"
        class="legend-toggle mb-2"
        mandatory
      >
        <v-btn value="source" size="x-small" class="legend-toggle-btn">
          <v-icon size="11" class="mr-1">mdi-database-outline</v-icon>
          Fuente
        </v-btn>
        <v-btn value="superCategory" size="x-small" class="legend-toggle-btn">
          <v-icon size="11" class="mr-1">mdi-shape-outline</v-icon>
          Tipo
        </v-btn>
      </v-btn-toggle>

      <div class="legend-divider" />

      <div v-for="item in legendItems" :key="item.key" class="legend-row">
        <span class="legend-dot" :style="{ background: item.color }" />
        <span class="legend-label">{{ item.label }}</span>
        <span class="legend-count">{{ item.count }}</span>
      </div>
    </div>

    <!-- Confidence detail dialog — lives outside Leaflet DOM -->
    <v-dialog v-model="showConfidenceDetail" max-width="420">
      <v-card v-if="detailAsset" rounded="lg">
        <v-card-title class="d-flex align-center pa-4 pb-2">
          <v-icon class="mr-2" :color="detailTierColor">mdi-shield-check</v-icon>
          Detalle de confianza
          <v-spacer />
          <v-btn icon size="small" variant="text" @click="showConfidenceDetail = false">
            <v-icon>mdi-close</v-icon>
          </v-btn>
        </v-card-title>

        <v-card-text class="pa-4 pt-0">
          <!-- Score summary -->
          <div class="score-summary d-flex align-center mb-4 pa-3 rounded-lg">
            <div>
              <div class="text-caption text-grey-darken-1">Score final</div>
              <div class="text-h5 font-weight-bold" :style="{ color: detailTierColor }">
                {{ (detailAsset.confidence_score * 100).toFixed(0) }}%
              </div>
            </div>
            <v-spacer />
            <v-chip :color="detailTierColor" variant="tonal" size="small">
              {{ detailTierLabel }}
            </v-chip>
          </div>

          <!-- Pipeline source -->
          <div class="text-caption text-grey-darken-1 mb-1">Fuente del análisis</div>
          <div class="d-flex ga-1 mb-4 flex-wrap">
            <v-chip
              v-for="src in detailAsset.data_sources"
              :key="src"
              size="x-small"
              variant="outlined"
              :color="detailSourceColor(src)"
            >
              <v-icon start size="10">{{ detailSourceIcon(src) }}</v-icon>
              {{ detailSourceLabel(src) }}
            </v-chip>
          </div>

          <!-- Signals breakdown -->
          <div v-if="detailSignalRows.length" class="mb-3">
            <div class="text-caption text-grey-darken-1 mb-2">Señales utilizadas</div>
            <div v-for="row in detailSignalRows" :key="row.key" class="mb-2">
              <div class="d-flex justify-space-between align-center mb-1">
                <div class="d-flex align-center">
                  <span class="text-caption font-weight-medium">{{ row.label }}</span>
                  <v-chip size="x-small" variant="tonal" color="grey" class="ml-1">
                    {{ (row.weight * 100).toFixed(0) }}%
                  </v-chip>
                </div>
                <span class="text-caption font-weight-bold" :style="{ color: signalColor(row.value) }">
                  {{ (row.value * 100).toFixed(0) }}%
                </span>
              </div>
              <v-progress-linear
                :model-value="row.value * 100"
                :color="signalColor(row.value)"
                height="5"
                rounded
                bg-color="grey-lighten-3"
              />
            </div>
          </div>

          <v-alert v-else type="info" variant="tonal" density="compact" class="text-caption">
            Detalle de señales no disponible para activos cargados desde caché.
          </v-alert>

          <div class="text-caption text-grey mt-3">
            <v-icon size="12" class="mr-1">mdi-information</v-icon>
            El score se suaviza con una distribución Beta para evitar valores extremos artificiales.
          </div>
        </v-card-text>
      </v-card>
    </v-dialog>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch, nextTick } from 'vue';
import { useAppStore } from '@/stores/store';
import type { Asset } from '@/types/types';
import {
  AssetCategory, CATEGORY_COLORS, CATEGORY_LABELS,
  AssetSuperCategory, SUPER_CATEGORY_COLORS, SUPER_CATEGORY_LABELS,
  CATEGORY_TO_SUPER,
} from '@/types/types';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import 'leaflet.markercluster';
import 'leaflet.markercluster/dist/MarkerCluster.css';
import 'leaflet.markercluster/dist/MarkerCluster.Default.css';

const store = useAppStore();
const mapContainer = ref<HTMLElement | null>(null);
const showConfidenceDetail = ref(false);
const detailAsset = ref<Asset | null>(null);
const colorMode = ref<'source' | 'superCategory'>('source');

let map: L.Map | null = null;
let markerClusterGroup: L.MarkerClusterGroup | null = null;
let markersById: Record<string, L.Marker> = {};

// --- Confidence detail dialog helpers ---
const MAPS_SIGNALS: Record<string, { label: string; weight: number }> = {
  name_match:        { label: 'Nombre incluye la empresa', weight: 0.30 },
  type_match:        { label: 'Tipo compatible con activo productivo', weight: 0.20 },
  address_corporate: { label: 'Señal de ubicación corporativa', weight: 0.15 },
  website_match:     { label: 'Dominio web corporativo', weight: 0.15 },
  reviews_b2b:       { label: 'Perfil de reseñas B2B', weight: 0.10 },
  llm_confidence:    { label: 'Valoración del modelo IA', weight: 0.10 },
};
const DOC_SIGNALS: Record<string, { label: string; weight: number }> = {
  evidence_strength:   { label: 'Fuerza de la evidencia documental', weight: 0.30 },
  address_specificity: { label: 'Especificidad de la dirección', weight: 0.20 },
  coordinate_source:   { label: 'Origen de las coordenadas', weight: 0.20 },
  name_quality:        { label: 'Calidad del nombre del activo', weight: 0.15 },
  llm_confidence:      { label: 'Valoración del modelo IA', weight: 0.15 },
};

const detailTierColor = computed(() => {
  if (!detailAsset.value) return 'grey';
  if (detailAsset.value.confidence_tier === 'HIGH') return 'green';
  if (detailAsset.value.confidence_tier === 'MEDIUM') return 'orange';
  return 'red';
});
const detailTierLabel = computed(() => {
  if (!detailAsset.value) return '';
  if (detailAsset.value.confidence_tier === 'HIGH') return 'Alta confianza';
  if (detailAsset.value.confidence_tier === 'MEDIUM') return 'Confianza media';
  return 'Baja confianza';
});
const detailSignalRows = computed(() => {
  const signals = detailAsset.value?.confidence_signals;
  if (!signals || Object.keys(signals).length === 0) return [];
  const meta = 'evidence_strength' in signals ? DOC_SIGNALS : MAPS_SIGNALS;
  return Object.entries(signals)
    .filter(([key]) => key in meta)
    .map(([key, value]) => ({ key, label: meta[key].label, weight: meta[key].weight, value }));
});

function signalColor(value: number): string {
  if (value >= 0.7) return 'green';
  if (value >= 0.4) return 'orange';
  return 'red';
}
function detailSourceLabel(src: string): string {
  const labels: Record<string, string> = { maps_api: 'Google Maps', document_upload: 'Documento', agent_search: 'Agente IA', llm_inference: 'Modelo IA' };
  return labels[src] ?? src;
}
function detailSourceIcon(src: string): string {
  const icons: Record<string, string> = { maps_api: 'mdi-google-maps', document_upload: 'mdi-file-document', agent_search: 'mdi-robot', llm_inference: 'mdi-brain' };
  return icons[src] ?? 'mdi-database';
}
function detailSourceColor(src: string): string {
  const colors: Record<string, string> = { maps_api: 'blue', document_upload: 'purple', agent_search: 'teal', llm_inference: 'grey' };
  return colors[src] ?? 'grey';
}

const emit = defineEmits<{
  (e: 'marker-click', asset: Asset): void;
}>();

const SPAIN_CENTER: L.LatLngExpression = [40.0, -3.7];
const SPAIN_ZOOM = 6;

const TILE_LAYERS: Record<string, { url: string; attribution: string }> = {
  Callejero: {
    url: 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  },
  Satélite: {
    url: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    attribution: '&copy; Esri',
  },
  Topográfico: {
    url: 'https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png',
    attribution: '&copy; OpenTopoMap',
  },
};

// Source color palette — distinct from category colors so the map legend is unambiguous.
const SOURCE_COLORS: Record<string, string> = {
  maps_api: '#2c6fff',        // blue
  document_upload: '#10b7c8', // teal
  agent_search: '#8b5cf6',    // purple
};
const SOURCE_LABELS: Record<string, string> = {
  maps_api: 'Maps API',
  document_upload: 'Document',
  agent_search: 'Agent Search',
};

function resolveSourceColor(dataSources: string[]): string {
  if (dataSources.includes('agent_search')) return SOURCE_COLORS.agent_search;
  if (dataSources.includes('document_upload')) return SOURCE_COLORS.document_upload;
  if (dataSources.includes('maps_api')) return SOURCE_COLORS.maps_api;
  return '#9E9E9E';
}

function resolveMarkerColor(asset: Asset): string {
  if (colorMode.value === 'superCategory') {
    const sc = CATEGORY_TO_SUPER[asset.category as AssetCategory];
    return sc ? SUPER_CATEGORY_COLORS[sc] : '#9E9E9E';
  }
  return resolveSourceColor(asset.data_sources || []);
}

const legendItems = computed(() => {
  if (colorMode.value === 'source') {
    return Object.entries(SOURCE_COLORS)
      .filter(([key]) => (store.sourceCounts[key] ?? 0) > 0)
      .map(([key, color]) => ({
        key,
        color,
        label: SOURCE_LABELS[key],
        count: store.sourceCounts[key] ?? 0,
      }));
  }
  return Object.values(AssetSuperCategory)
    .filter((sc) => (store.superCategoryCounts[sc] ?? 0) > 0)
    .map((sc) => ({
      key: sc,
      color: SUPER_CATEGORY_COLORS[sc],
      label: SUPER_CATEGORY_LABELS[sc],
      count: store.superCategoryCounts[sc] ?? 0,
    }));
});

function createMarkerIcon(color: string, tier: string): L.DivIcon {
  const opacity = tier === 'LOW' ? '0.55' : '1';
  const size = tier === 'HIGH' ? 14 : 10;
  return L.divIcon({
    className: 'custom-marker',
    html: `<div style="
      width:${size}px;height:${size}px;border-radius:50%;
      background:${color};opacity:${opacity};
      border:2px solid white;box-shadow:0 1px 4px rgba(0,0,0,0.4);
    "></div>`,
    iconSize: [size + 4, size + 4],
    iconAnchor: [(size + 4) / 2, (size + 4) / 2],
  });
}

function buildPopupContent(asset: Asset): string {
  const catLabel = CATEGORY_LABELS[asset.category as AssetCategory] || asset.category;
  const color = CATEGORY_COLORS[asset.category as AssetCategory] || '#9E9E9E';
  const confPct = (asset.confidence_score * 100).toFixed(0);
  const tierColor =
    asset.confidence_tier === 'HIGH' ? '#4CAF50' : asset.confidence_tier === 'MEDIUM' ? '#FF9800' : '#F44336';
  const tags = (asset.functional_tags || [])
    .slice(0, 4)
    .map(
      (t) =>
        `<span style="display:inline-block;padding:2px 8px;margin:2px;border:1px solid rgba(88,114,160,0.3);border-radius:999px;font-size:10px;color:#4f648e;background:rgba(236,243,255,0.9);">${t}</span>`,
    )
    .join('');
  const hqBadge = asset.is_headquarters
    ? '<span style="display:inline-block;padding:2px 8px;margin-left:6px;background:#fef4cc;color:#8c6a00;border-radius:999px;font-size:10px;font-weight:600;">HQ</span>'
    : '';

  const sourceChipStyles: Record<string, string> = {
    maps_api:        'background:rgba(44,111,255,0.12);color:#2c6fff;',
    document_upload: 'background:rgba(16,183,200,0.12);color:#0e8e9c;',
    agent_search:    'background:rgba(139,92,246,0.12);color:#8b5cf6;',
  };
  const sourceChips = (asset.data_sources || [])
    .filter((src: string) => !src.includes('inference'))
    .map((src: string) => {
      const style = sourceChipStyles[src] || 'background:rgba(100,100,100,0.1);color:#666;';
      const label = SOURCE_LABELS[src] || src;
      return `<span style="display:inline-block;padding:2px 8px;margin:2px;${style}border-radius:999px;font-size:9px;font-weight:600;">${label}</span>`;
    })
    .join('');

  return `
    <div style="max-width:272px;font-family:'Sora',sans-serif;font-size:13px;color:#1b2e55;">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px;">
        <div style="font-weight:700;line-height:1.3;">${asset.name}</div>
        ${hqBadge}
      </div>
      <div style="display:flex;flex-wrap:wrap;align-items:center;gap:4px;margin-bottom:8px;">
        <span style="display:inline-block;padding:2px 9px;background:${color};color:white;border-radius:999px;font-size:10px;font-weight:600;">${catLabel}</span>
        ${sourceChips}
      </div>
      <div style="color:#5f739f;font-size:11px;margin-bottom:6px;line-height:1.45;">${asset.address}</div>
      ${asset.description ? `<div style="font-size:11px;margin-bottom:6px;color:#42567f;line-height:1.4;">${asset.description}</div>` : ''}
      <div style="display:flex;align-items:center;margin-bottom:8px;">
        <span style="font-size:11px;margin-right:6px;color:#54668f;">Confianza:</span>
        <div style="flex:1;height:7px;background:#e7eefb;border-radius:6px;max-width:110px;">
          <div style="height:7px;border-radius:6px;background:${tierColor};width:${confPct}%;"></div>
        </div>
        <span style="font-size:11px;margin-left:6px;font-weight:700;color:#1f3460;">${confPct}%</span>
        <button
          data-action="confidence-detail"
          data-asset-id="${asset.id}"
          style="margin-left:6px;background:none;border:none;cursor:pointer;padding:0;color:${tierColor};display:flex;align-items:center;"
          title="Ver cálculo de confianza"
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
            <path d="M13 9h-2V7h2m0 10h-2v-6h2m-1-9A10 10 0 0 0 2 12a10 10 0 0 0 10 10 10 10 0 0 0 10-10A10 10 0 0 0 12 2z"/>
          </svg>
        </button>
      </div>
      ${tags ? `<div style="margin-bottom:4px;">${tags}</div>` : ''}
      <a href="https://www.google.com/maps/place/?q=place_id:${asset.google_place_id}" target="_blank" style="font-size:11px;color:#2b61da;text-decoration:none;font-weight:600;">Ver en Google Maps ↗</a>
    </div>
  `;
}

function addMarkers() {
  if (!map) return;

  if (markerClusterGroup) {
    map.removeLayer(markerClusterGroup);
  }

  markerClusterGroup = L.markerClusterGroup({
    disableClusteringAtZoom: store.clusteringEnabled ? undefined : 1,
    maxClusterRadius: store.clusteringEnabled ? 50 : 0,
    spiderfyOnMaxZoom: true,
    showCoverageOnHover: false,
  });

  markersById = {};
  const assets = store.filteredAssets;

  for (const asset of assets) {
    const marker = L.marker([asset.latitude, asset.longitude], {
      icon: createMarkerIcon(resolveMarkerColor(asset), asset.confidence_tier),
    });
    marker.bindPopup(buildPopupContent(asset), { maxWidth: 280 });
    marker.on('click', () => {
      store.selectAsset(asset.id);
      emit('marker-click', asset);
    });
    markerClusterGroup.addLayer(marker);
    markersById[asset.id] = marker;
  }

  map.addLayer(markerClusterGroup);

  if (assets.length > 0) {
    const group = L.featureGroup(Object.values(markersById));
    map.fitBounds(group.getBounds().pad(0.1));
  }
}

function flyToAsset(asset: Asset) {
  if (!map) return;
  map.flyTo([asset.latitude, asset.longitude], 14, { duration: 1 });
  const marker = markersById[asset.id];
  if (marker) {
    setTimeout(() => marker.openPopup(), 500);
  }
}

defineExpose({ flyToAsset });

onMounted(() => {
  if (!mapContainer.value) return;

  map = L.map(mapContainer.value, {
    center: SPAIN_CENTER,
    zoom: SPAIN_ZOOM,
    zoomControl: true,
  });

  const baseLayers: Record<string, L.TileLayer> = {};
  let first = true;
  for (const [name, cfg] of Object.entries(TILE_LAYERS)) {
    const layer = L.tileLayer(cfg.url, { attribution: cfg.attribution });
    baseLayers[name] = layer;
    if (first) {
      layer.addTo(map);
      first = false;
    }
  }

  L.control.layers(baseLayers).addTo(map);

  nextTick(() => addMarkers());

  mapContainer.value.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest('[data-action="confidence-detail"]');
    if (!btn) return;
    e.stopPropagation();
    const assetId = btn.getAttribute('data-asset-id');
    const asset = store.assets.find((a) => a.id === assetId) ?? null;
    detailAsset.value = asset;
    showConfidenceDetail.value = true;
  });
});

watch(
  () => [store.filteredAssets, store.clusteringEnabled],
  () => { addMarkers(); },
  { deep: true },
);

watch(colorMode, () => { addMarkers(); });

watch(
  () => store.selectedAssetId,
  (newId) => {
    if (newId) {
      const asset = store.assets.find((a) => a.id === newId);
      if (asset) flyToAsset(asset);
    }
  },
);

onUnmounted(() => {
  if (map) {
    map.remove();
    map = null;
  }
});
</script>

<style>
.map-container {
  z-index: 0;
}

.custom-marker {
  background: transparent !important;
  border: none !important;
}

.leaflet-popup-content-wrapper,
.leaflet-popup-tip {
  background: rgba(255, 255, 255, 0.96);
  border: 1px solid rgba(116, 145, 196, 0.32);
  box-shadow: 0 16px 36px rgba(22, 57, 124, 0.16);
}

.leaflet-popup-content {
  margin: 12px 12px 10px;
}

.leaflet-container a.leaflet-popup-close-button {
  color: #4d6696;
}

.leaflet-control-zoom,
.leaflet-control-layers {
  border: 1px solid rgba(108, 141, 200, 0.28) !important;
  border-radius: 12px !important;
  overflow: hidden;
  box-shadow: 0 14px 26px rgba(19, 52, 118, 0.12);
}

.leaflet-control-zoom a,
.leaflet-control-layers-toggle,
.leaflet-control-layers-expanded {
  background: rgba(255, 255, 255, 0.94) !important;
  color: #223a69 !important;
}

.score-summary {
  background: rgba(0, 0, 0, 0.04);
}

/* ── Legend overlay ──────────────────────────────────────────── */
.legend-overlay {
  position: absolute;
  bottom: 28px;
  right: 10px;
  z-index: 800;
  background: rgba(255, 255, 255, 0.95);
  border: 1px solid rgba(108, 141, 200, 0.28);
  border-radius: 14px;
  padding: 0.6rem 0.8rem 0.55rem;
  box-shadow: 0 8px 22px rgba(19, 52, 118, 0.12);
  font-family: 'Sora', sans-serif;
  font-size: 11px;
  color: #223a69;
  min-width: 140px;
  backdrop-filter: blur(6px);
}

.legend-toggle {
  width: 100%;
  border-radius: 8px !important;
  overflow: hidden;
}

.legend-toggle-btn {
  flex: 1;
  font-size: 10px !important;
  letter-spacing: 0.02em;
  text-transform: none !important;
}

.legend-divider {
  height: 1px;
  background: rgba(108, 141, 200, 0.2);
  margin: 0.45rem 0;
}

.legend-row {
  display: flex;
  align-items: center;
  gap: 7px;
  padding: 2px 0;
}

.legend-dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
  flex-shrink: 0;
  border: 1.5px solid white;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.3);
}

.legend-label {
  flex: 1;
  font-size: 11px;
  color: #223a69;
}

.legend-count {
  font-size: 10px;
  color: #7894c4;
  font-weight: 600;
}
</style>
