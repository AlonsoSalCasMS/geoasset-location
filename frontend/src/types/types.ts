export interface Company {
  id: string
  name: string
  address?: string
  types?: string[]
  website?: string
  cif?: string
  sector?: string
  cnae?: string
  logo_url?: string
  headquarters?: string
}

// ── Supercategory (nivel superior de clasificación) ─────────────
export enum AssetSuperCategory {
  CORP  = 'CORP',
  OPS   = 'OPS',
  COM   = 'COM',
  INFRA = 'INFRA',
  OTR   = 'OTR',
}

export const SUPER_CATEGORY_LABELS: Record<AssetSuperCategory, string> = {
  [AssetSuperCategory.CORP]:  'Corporativo',
  [AssetSuperCategory.OPS]:   'Operaciones',
  [AssetSuperCategory.COM]:   'Comercial / Servicios',
  [AssetSuperCategory.INFRA]: 'Infraestructura',
  [AssetSuperCategory.OTR]:   'Otro',
}

export const SUPER_CATEGORY_ICONS: Record<AssetSuperCategory, string> = {
  [AssetSuperCategory.CORP]:  'mdi-office-building-outline',
  [AssetSuperCategory.OPS]:   'mdi-cogs',
  [AssetSuperCategory.COM]:   'mdi-shopping-outline',
  [AssetSuperCategory.INFRA]: 'mdi-transmission-tower',
  [AssetSuperCategory.OTR]:   'mdi-map-marker-outline',
}

export const SUPER_CATEGORY_COLORS: Record<AssetSuperCategory, string> = {
  [AssetSuperCategory.CORP]:  '#1565C0',
  [AssetSuperCategory.OPS]:   '#E65100',
  [AssetSuperCategory.COM]:   '#2E7D32',
  [AssetSuperCategory.INFRA]: '#546E7A',
  [AssetSuperCategory.OTR]:   '#757575',
}

// ── Subcategory (clasificación detallada existente) ─────────────
export enum AssetCategory {
  HQ = 'HQ',
  OFF = 'OFF',
  FAB = 'FAB',
  LOG = 'LOG',
  TEC = 'TEC',
  COM = 'COM',
  AGR = 'AGR',
  ENE = 'ENE',
  TRA = 'TRA',
  HOT = 'HOT',
  SAN = 'SAN',
  OTR = 'OTR',
}

export const CATEGORY_LABELS: Record<AssetCategory, string> = {
  [AssetCategory.HQ]: 'Sede Central',
  [AssetCategory.OFF]: 'Oficina Regional',
  [AssetCategory.FAB]: 'Fábrica / Planta',
  [AssetCategory.LOG]: 'Centro Logístico',
  [AssetCategory.TEC]: 'Centro Tecnológico',
  [AssetCategory.COM]: 'Punto de Venta',
  [AssetCategory.AGR]: 'Explotación Agrícola',
  [AssetCategory.ENE]: 'Inst. Energética',
  [AssetCategory.TRA]: 'Infraestructura Transporte',
  [AssetCategory.HOT]: 'Activo Hotelero',
  [AssetCategory.SAN]: 'Centro Sanitario',
  [AssetCategory.OTR]: 'Otro',
}

export const CATEGORY_ICONS: Record<AssetCategory, string> = {
  [AssetCategory.HQ]: 'mdi-office-building',
  [AssetCategory.OFF]: 'mdi-domain',
  [AssetCategory.FAB]: 'mdi-factory',
  [AssetCategory.LOG]: 'mdi-warehouse',
  [AssetCategory.TEC]: 'mdi-flask',
  [AssetCategory.COM]: 'mdi-store',
  [AssetCategory.AGR]: 'mdi-sprout',
  [AssetCategory.ENE]: 'mdi-lightning-bolt',
  [AssetCategory.TRA]: 'mdi-train',
  [AssetCategory.HOT]: 'mdi-bed',
  [AssetCategory.SAN]: 'mdi-hospital-box',
  [AssetCategory.OTR]: 'mdi-map-marker',
}

export const CATEGORY_COLORS: Record<AssetCategory, string> = {
  [AssetCategory.HQ]: '#1565C0',
  [AssetCategory.OFF]: '#2196F3',
  [AssetCategory.FAB]: '#E65100',
  [AssetCategory.LOG]: '#FF9800',
  [AssetCategory.TEC]: '#7B1FA2',
  [AssetCategory.COM]: '#4CAF50',
  [AssetCategory.AGR]: '#33691E',
  [AssetCategory.ENE]: '#FDD835',
  [AssetCategory.TRA]: '#607D8B',
  [AssetCategory.HOT]: '#E91E63',
  [AssetCategory.SAN]: '#F44336',
  [AssetCategory.OTR]: '#9E9E9E',
}

// ── Mapping between the two levels ─────────────────────────────
export const CATEGORY_TO_SUPER: Record<AssetCategory, AssetSuperCategory> = {
  [AssetCategory.HQ]:  AssetSuperCategory.CORP,
  [AssetCategory.OFF]: AssetSuperCategory.CORP,
  [AssetCategory.TEC]: AssetSuperCategory.CORP,
  [AssetCategory.FAB]: AssetSuperCategory.OPS,
  [AssetCategory.LOG]: AssetSuperCategory.OPS,
  [AssetCategory.AGR]: AssetSuperCategory.OPS,
  [AssetCategory.COM]: AssetSuperCategory.COM,
  [AssetCategory.HOT]: AssetSuperCategory.COM,
  [AssetCategory.SAN]: AssetSuperCategory.COM,
  [AssetCategory.ENE]: AssetSuperCategory.INFRA,
  [AssetCategory.TRA]: AssetSuperCategory.INFRA,
  [AssetCategory.OTR]: AssetSuperCategory.OTR,
}

export const SUPER_TO_CATEGORIES: Record<AssetSuperCategory, AssetCategory[]> = {
  [AssetSuperCategory.CORP]:  [AssetCategory.HQ, AssetCategory.OFF, AssetCategory.TEC],
  [AssetSuperCategory.OPS]:   [AssetCategory.FAB, AssetCategory.LOG, AssetCategory.AGR],
  [AssetSuperCategory.COM]:   [AssetCategory.COM, AssetCategory.HOT, AssetCategory.SAN],
  [AssetSuperCategory.INFRA]: [AssetCategory.ENE, AssetCategory.TRA],
  [AssetSuperCategory.OTR]:   [AssetCategory.OTR],
}

export interface Asset {
  id: string
  company_id: string
  name: string
  raw_name: string
  super_category?: AssetSuperCategory
  category: AssetCategory
  subcategory?: string
  latitude: number
  longitude: number
  address: string
  municipality: string
  province: string
  autonomous_community: string
  postal_code?: string
  description?: string
  size_estimate?: string
  functional_tags: string[]
  is_headquarters: boolean
  google_place_id: string
  confidence_score: number
  confidence_tier: string
  confidence_signals?: Record<string, number>
  data_sources: string[]
  website?: string
  phone?: string
  created_at?: string
  updated_at?: string
}

export interface PipelineStep {
  step: number
  name: string
  status: 'pending' | 'running' | 'complete' | 'error'
  estimated_seconds?: number
  found?: number
  error?: string
}

export interface AnalysisMetadata {
  company?: Company
  total_assets: number
  high_confidence?: number
  medium_confidence?: number
  low_confidence?: number
  last_updated?: string
}

export type AssetSource = 'maps_api' | 'document_upload' | 'agent_search'

export type AnalysisMode = 'search' | 'document' | 'combined'

export type AppView = 'search' | 'agent' | 'agent_review' | 'processing' | 'results'

export interface AgentFile {
  filename: string
  size: number
  url: string
  relevance_reason: string
  session_id: string
  extension?: string
  page_count?: number | null
}

export type AgentEventType =
  | 'thinking'
  | 'searching'
  | 'found_urls'
  | 'downloading'
  | 'accepted'
  | 'rejected'
  | 'error'

export interface AgentEvent {
  type: AgentEventType
  content: string
  filename?: string
  url?: string
  timestamp: number
}
