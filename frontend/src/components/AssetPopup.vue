<template>
  <div v-if="asset" class="asset-popup">
    <div class="d-flex align-center mb-2">
      <v-icon :color="categoryColor" size="20" class="mr-2">{{ categoryIcon }}</v-icon>
      <span class="text-subtitle-2 font-weight-bold">{{ asset.name }}</span>
    </div>

    <v-chip :color="superCategoryColor" size="x-small" variant="flat" class="mb-1 mr-1">
      <v-icon start size="11">{{ superCategoryIcon }}</v-icon>
      {{ superCategoryLabel }}
    </v-chip>
    <v-chip :color="categoryColor" size="x-small" variant="tonal" class="mb-2">
      {{ categoryLabel }}
    </v-chip>
    <v-chip
      v-if="asset.is_headquarters"
      size="x-small"
      color="amber"
      variant="tonal"
      class="mb-2 ml-1"
    >
      HQ
    </v-chip>

    <div class="text-caption text-grey-darken-1 mb-1">
      <v-icon size="12" class="mr-1">mdi-map-marker</v-icon>{{ asset.address }}
    </div>

    <div v-if="asset.description" class="text-caption mb-2">{{ asset.description }}</div>

    <!-- Confidence bar -->
    <div class="d-flex align-center mb-2">
      <span class="text-caption mr-2">Confianza:</span>
      <v-progress-linear
        :model-value="asset.confidence_score * 100"
        :color="tierColor"
        height="6"
        rounded
        style="max-width: 100px"
      />
      <span class="text-caption ml-2 font-weight-medium">{{ (asset.confidence_score * 100).toFixed(0) }}%</span>
      <v-btn
        icon
        size="x-small"
        variant="text"
        :color="tierColor"
        class="ml-1"
        @click.stop="showDetail = true"
      >
        <v-icon size="14">mdi-information-outline</v-icon>
        <v-tooltip activator="parent" location="top">Ver cálculo de confianza</v-tooltip>
      </v-btn>
    </div>

    <!-- Tags -->
    <div v-if="asset.functional_tags?.length" class="mb-2">
      <v-chip
        v-for="tag in asset.functional_tags.slice(0, 4)"
        :key="tag"
        size="x-small"
        variant="outlined"
        class="mr-1 mb-1"
      >
        {{ tag }}
      </v-chip>
    </div>

    <!-- Google Maps link -->
    <a
      :href="`https://www.google.com/maps/place/?q=place_id:${asset.google_place_id}`"
      target="_blank"
      class="text-caption text-decoration-none"
    >
      <v-icon size="12" class="mr-1">mdi-open-in-new</v-icon>Ver en Google Maps
    </a>

    <!-- Confidence detail dialog -->
    <v-dialog v-model="showDetail" max-width="420" @click:outside="showDetail = false">
      <v-card rounded="lg">
        <v-card-title class="d-flex align-center pa-4 pb-2">
          <v-icon class="mr-2" :color="tierColor">mdi-shield-check</v-icon>
          Detalle de confianza
          <v-spacer />
          <v-btn icon size="small" variant="text" @click="showDetail = false">
            <v-icon>mdi-close</v-icon>
          </v-btn>
        </v-card-title>

        <v-card-text class="pa-4 pt-0">
          <!-- Score summary -->
          <div class="score-summary d-flex align-center mb-4 pa-3 rounded-lg">
            <div>
              <div class="text-caption text-grey-darken-1">Score final</div>
              <div class="text-h5 font-weight-bold" :style="{ color: tierColor }">
                {{ (asset.confidence_score * 100).toFixed(0) }}%
              </div>
            </div>
            <v-spacer />
            <v-chip :color="tierColor" variant="tonal" size="small">
              {{ tierLabel }}
            </v-chip>
          </div>

          <!-- Pipeline source -->
          <div class="text-caption text-grey-darken-1 mb-1">Fuente del análisis</div>
          <div class="d-flex ga-1 mb-4 flex-wrap">
            <v-chip
              v-for="src in asset.data_sources"
              :key="src"
              size="x-small"
              variant="outlined"
              :color="sourceColor(src)"
            >
              <v-icon start size="10">{{ sourceIcon(src) }}</v-icon>
              {{ sourceLabel(src) }}
            </v-chip>
          </div>

          <!-- Signals breakdown -->
          <div v-if="signalRows.length" class="mb-3">
            <div class="text-caption text-grey-darken-1 mb-2">Señales utilizadas</div>
            <div v-for="row in signalRows" :key="row.key" class="signal-row mb-2">
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

          <!-- Beta smoothing note -->
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
import { ref, computed } from 'vue';
import type { Asset } from '@/types/types';
import {
  AssetCategory, CATEGORY_LABELS, CATEGORY_ICONS, CATEGORY_COLORS,
  CATEGORY_TO_SUPER, SUPER_CATEGORY_LABELS, SUPER_CATEGORY_ICONS, SUPER_CATEGORY_COLORS,
} from '@/types/types';

const props = defineProps<{ asset: Asset }>();

const showDetail = ref(false);

const categoryLabel = computed(() => CATEGORY_LABELS[props.asset.category as AssetCategory] || props.asset.category);
const categoryIcon = computed(() => CATEGORY_ICONS[props.asset.category as AssetCategory] || 'mdi-map-marker');
const categoryColor = computed(() => CATEGORY_COLORS[props.asset.category as AssetCategory] || '#9E9E9E');

const superCat = computed(() => CATEGORY_TO_SUPER[props.asset.category as AssetCategory]);
const superCategoryLabel = computed(() => superCat.value ? SUPER_CATEGORY_LABELS[superCat.value] : '');
const superCategoryIcon = computed(() => superCat.value ? SUPER_CATEGORY_ICONS[superCat.value] : 'mdi-map-marker-outline');
const superCategoryColor = computed(() => superCat.value ? SUPER_CATEGORY_COLORS[superCat.value] : '#757575');

const tierColor = computed(() => {
  if (props.asset.confidence_tier === 'HIGH') return 'green';
  if (props.asset.confidence_tier === 'MEDIUM') return 'orange';
  return 'red';
});

const tierLabel = computed(() => {
  if (props.asset.confidence_tier === 'HIGH') return 'Alta confianza';
  if (props.asset.confidence_tier === 'MEDIUM') return 'Confianza media';
  return 'Baja confianza';
});

// Signal metadata per pipeline type
const MAPS_SIGNALS: Record<string, { label: string; weight: number }> = {
  name_match: { label: 'Nombre incluye la empresa', weight: 0.30 },
  type_match: { label: 'Tipo compatible con activo productivo', weight: 0.20 },
  address_corporate: { label: 'Señal de ubicación corporativa', weight: 0.15 },
  website_match: { label: 'Dominio web corporativo', weight: 0.15 },
  reviews_b2b: { label: 'Perfil de reseñas B2B', weight: 0.10 },
  llm_confidence: { label: 'Valoración del modelo IA', weight: 0.10 },
};

const DOC_SIGNALS: Record<string, { label: string; weight: number }> = {
  evidence_strength: { label: 'Fuerza de la evidencia documental', weight: 0.30 },
  address_specificity: { label: 'Especificidad de la dirección', weight: 0.20 },
  coordinate_source: { label: 'Origen de las coordenadas', weight: 0.20 },
  name_quality: { label: 'Calidad del nombre del activo', weight: 0.15 },
  llm_confidence: { label: 'Valoración del modelo IA', weight: 0.15 },
};

const signalRows = computed(() => {
  const signals = props.asset.confidence_signals;
  if (!signals || Object.keys(signals).length === 0) return [];

  const meta = 'evidence_strength' in signals ? DOC_SIGNALS : MAPS_SIGNALS;

  return Object.entries(signals)
    .filter(([key]) => key in meta)
    .map(([key, value]) => ({
      key,
      label: meta[key].label,
      weight: meta[key].weight,
      value,
    }));
});

function signalColor(value: number): string {
  if (value >= 0.7) return 'green';
  if (value >= 0.4) return 'orange';
  return 'red';
}

function sourceLabel(src: string): string {
  if (src === 'maps_api') return 'Google Maps';
  if (src === 'document_upload') return 'Documento';
  if (src === 'agent_search') return 'Agente IA';
  if (src === 'llm_inference') return 'Modelo IA';
  return src;
}

function sourceIcon(src: string): string {
  if (src === 'maps_api') return 'mdi-google-maps';
  if (src === 'document_upload') return 'mdi-file-document';
  if (src === 'agent_search') return 'mdi-robot';
  if (src === 'llm_inference') return 'mdi-brain';
  return 'mdi-database';
}

function sourceColor(src: string): string {
  if (src === 'maps_api') return 'blue';
  if (src === 'document_upload') return 'purple';
  if (src === 'agent_search') return 'teal';
  if (src === 'llm_inference') return 'grey';
  return 'grey';
}
</script>

<style scoped>
.asset-popup {
  max-width: 280px;
  font-family: 'Montserrat', sans-serif;
}

.score-summary {
  background: rgba(0, 0, 0, 0.04);
}
</style>
