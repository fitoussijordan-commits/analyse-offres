// Client Supabase léger (REST API, pas de dépendance externe)
const SUPABASE_URL = 'https://fcjtntvuuhmrqgafdsjl.supabase.co/rest/v1'
const SUPABASE_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZjanRudHZ1dWhtcnFnYWZkc2psIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ1MTI2OTYsImV4cCI6MjA5MDA4ODY5Nn0.dx8b_rkv7Lt-9K-xGq9-z9OnLsolFNnWJfoTTA8re7M'

const H: Record<string, string> = {
  apikey: SUPABASE_KEY,
  Authorization: `Bearer ${SUPABASE_KEY}`,
  'Content-Type': 'application/json',
}

async function req<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${SUPABASE_URL}${path}`, {
    ...init,
    headers: { ...H, ...(init?.headers as Record<string, string> | undefined) },
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`Supabase ${res.status}: ${text}`)
  }
  if (res.status === 204) return [] as unknown as T
  return res.json()
}

// ─── Types ────────────────────────────────────────────────────────────────────

export interface Produit {
  ref: string
  material_text: string
  material_id: string | null
  product_range: string
  package_size: string
  content: string
  tray_quantity: number
  conso_moyenne: number
  actif: boolean
}

export interface Quantite {
  id: string
  ref: string
  mois: string // format YYYY-MM
  quantite: number
}

export interface HistEntry {
  id: string
  ref: string
  produit_label: string
  mois: string
  ancienne_qte: number | null
  nouvelle_qte: number
  variation: number
  raison: string
  modifie_par: string
  modifie_le: string
}

// ─── API ──────────────────────────────────────────────────────────────────────

export const getProduits = () =>
  req<Produit[]>(
    '/planning_produits?order=product_range.asc,material_text.asc&actif=eq.true'
  )

export const getQuantites = (year: number) =>
  req<Quantite[]>(
    `/planning_quantites?mois=gte.${year}-01&mois=lte.${year}-12`
  )

export const upsertQuantite = (data: {
  ref: string
  mois: string
  quantite: number
}) =>
  req<Quantite[]>('/planning_quantites?on_conflict=ref,mois', {
    method: 'POST',
    headers: { Prefer: 'resolution=merge-duplicates,return=representation' },
    body: JSON.stringify(data),
  })

export const addHistorique = (data: {
  ref: string
  produit_label: string
  mois: string
  ancienne_qte: number
  nouvelle_qte: number
  raison: string
}) =>
  req('/planning_historique', {
    method: 'POST',
    headers: { Prefer: 'return=minimal' },
    body: JSON.stringify(data),
  })

export const getHistorique = (limit = 100) =>
  req<HistEntry[]>(
    `/planning_historique?order=modifie_le.desc&limit=${limit}`
  )

export const insertProduits = (produits: Omit<Produit, 'actif'>[]) =>
  req('/planning_produits?on_conflict=ref', {
    method: 'POST',
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify(produits),
  })

export const insertQuantites = (quantites: { ref: string; mois: string; quantite: number }[]) =>
  req('/planning_quantites?on_conflict=ref,mois', {
    method: 'POST',
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify(quantites),
  })
