"use client"
import { useState, useEffect, useMemo, useRef } from "react"
import * as sb from "@/lib/supabase-planning"

// ── Couleurs (cohérentes avec le reste de l'app) ──────────────────────────────
const C = {
  bg: "#f8fafc", white: "#ffffff", text: "#1a1a2e", textSec: "#374151",
  textMuted: "#6b7280", border: "#e5e7eb",
  blue: "#3b82f6", blueSoft: "#eff6ff",
  green: "#22c55e", greenSoft: "#f0fdf4",
  red: "#ef4444", redSoft: "#fef2f2",
  orange: "#f97316", orangeSoft: "#fff7ed",
  purple: "#7c3aed", purpleSoft: "#f5f3ff",
  teal: "#0d9488", tealSoft: "#f0fdfa",
  shadow: "0 1px 4px rgba(0,0,0,0.07)",
  shadowMd: "0 4px 16px rgba(0,0,0,0.12)",
}

const MONTH_SHORT = ['Jan','Fév','Mar','Avr','Mai','Jun','Jul','Aoû','Sep','Oct','Nov','Déc']
const MONTH_FR    = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre']

function moisKey(year: number, idx: number) {
  return `${year}-${String(idx + 1).padStart(2, '0')}`
}
function moisLabel(mois: string) {
  const [y, m] = mois.split('-')
  return `${MONTH_FR[parseInt(m) - 1]} ${y}`
}
function fmt(n: number) {
  return n === 0 ? '—' : n.toLocaleString('fr-FR')
}

// ── Types ─────────────────────────────────────────────────────────────────────
interface EditCell {
  ref: string
  mois: string
  productLabel: string
  oldQte: number
}

interface Props {
  onToast: (msg: string, type?: 'success' | 'error' | 'info') => void
}

// ── Styles partagés ───────────────────────────────────────────────────────────
const thBase: React.CSSProperties = {
  padding: '9px 8px',
  fontSize: 11,
  fontWeight: 700,
  color: C.textMuted,
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
  borderRight: `1px solid ${C.border}`,
  whiteSpace: 'nowrap',
  background: C.bg,
  userSelect: 'none',
}

const tdBase: React.CSSProperties = {
  padding: '7px 8px',
  fontSize: 13,
  borderRight: `1px solid ${C.border}`,
  verticalAlign: 'middle',
}

// ── Composant principal ───────────────────────────────────────────────────────
export default function PlanningTab({ onToast }: Props) {
  const [produits, setProduits]   = useState<sb.Produit[]>([])
  const [qtyMap, setQtyMap]       = useState<Record<string, Record<string, number>>>({})
  const [loading, setLoading]     = useState(true)
  const [year, setYear]           = useState(2026)
  const [search, setSearch]       = useState('')
  const [rangeFilter, setRange]   = useState('')
  const [view, setView]           = useState<'grid' | 'historique'>('grid')

  // Edition
  const [editCell, setEditCell]   = useState<EditCell | null>(null)
  const [editQte, setEditQte]     = useState('')
  const [editRaison, setEditRaison] = useState('')
  const [saving, setSaving]       = useState(false)
  const editRef                   = useRef<HTMLInputElement>(null)

  // Historique
  const [historique, setHistorique] = useState<sb.HistEntry[]>([])
  const [histLoading, setHistLoading] = useState(false)
  const [histSearch, setHistSearch] = useState('')

  // ── Chargement données ──────────────────────────────────────────────────────
  useEffect(() => {
    setLoading(true)
    Promise.all([sb.getProduits(), sb.getQuantites(year)])
      .then(([prods, qtys]) => {
        setProduits(prods)
        const map: Record<string, Record<string, number>> = {}
        for (const q of qtys) {
          if (!map[q.ref]) map[q.ref] = {}
          map[q.ref][q.mois] = q.quantite
        }
        setQtyMap(map)
      })
      .catch(e => onToast('Erreur chargement: ' + e.message, 'error'))
      .finally(() => setLoading(false))
  }, [year])

  useEffect(() => {
    if (view !== 'historique') return
    setHistLoading(true)
    sb.getHistorique(200)
      .then(setHistorique)
      .catch(e => onToast('Erreur historique: ' + e.message, 'error'))
      .finally(() => setHistLoading(false))
  }, [view])

  useEffect(() => {
    if (editCell) setTimeout(() => editRef.current?.focus(), 50)
  }, [editCell])

  // ── Dérivés ─────────────────────────────────────────────────────────────────
  const ranges = useMemo(
    () => [...new Set(produits.map(p => p.product_range).filter(Boolean))].sort(),
    [produits]
  )

  const filtered = useMemo(() => {
    const s = search.toLowerCase()
    return produits.filter(p =>
      (!rangeFilter || p.product_range === rangeFilter) &&
      (!s || String(p.ref).toLowerCase().includes(s) || p.material_text.toLowerCase().includes(s))
    )
  }, [produits, search, rangeFilter])

  const months = useMemo(() => Array.from({ length: 12 }, (_, i) => moisKey(year, i)), [year])

  const monthTotals = useMemo(
    () => months.map(m => filtered.reduce((s, p) => s + (qtyMap[p.ref]?.[m] ?? 0), 0)),
    [filtered, qtyMap, months]
  )

  const grandTotal = monthTotals.reduce((a, b) => a + b, 0)

  const filteredHist = useMemo(() => {
    const s = histSearch.toLowerCase()
    return s
      ? historique.filter(h => h.ref.toLowerCase().includes(s) || h.produit_label.toLowerCase().includes(s) || h.raison.toLowerCase().includes(s))
      : historique
  }, [historique, histSearch])

  // ── Edition ─────────────────────────────────────────────────────────────────
  const openEdit = (ref: string, mois: string, productLabel: string) => {
    const oldQte = qtyMap[ref]?.[mois] ?? 0
    setEditCell({ ref, mois, productLabel, oldQte })
    setEditQte(String(oldQte))
    setEditRaison('')
  }

  const saveEdit = async () => {
    if (!editCell) return
    const newQte = Math.max(0, parseInt(editQte) || 0)
    if (newQte === editCell.oldQte) { setEditCell(null); return }
    setSaving(true)
    try {
      await sb.upsertQuantite({ ref: editCell.ref, mois: editCell.mois, quantite: newQte })
      await sb.addHistorique({
        ref: editCell.ref,
        produit_label: editCell.productLabel,
        mois: editCell.mois,
        ancienne_qte: editCell.oldQte,
        nouvelle_qte: newQte,
        raison: editRaison.trim(),
      })
      setQtyMap(prev => ({
        ...prev,
        [editCell.ref]: { ...(prev[editCell.ref] ?? {}), [editCell.mois]: newQte },
      }))
      const diff = newQte - editCell.oldQte
      onToast(
        `${editCell.ref} · ${MONTH_SHORT[parseInt(editCell.mois.split('-')[1]) - 1]} : ${editCell.oldQte} → ${newQte} (${diff > 0 ? '+' : ''}${diff})`,
        'success'
      )
      setEditCell(null)
    } catch (e: any) {
      onToast('Erreur: ' + e.message, 'error')
    } finally {
      setSaving(false)
    }
  }

  // ── Export CSV ───────────────────────────────────────────────────────────────
  const exportCSV = () => {
    const rows = [
      ['REF', 'Produit', 'Gamme', ...MONTH_SHORT.map(m => `${m} ${year}`), 'Total'],
      ...filtered.map(p => {
        const qtys = months.map(m => qtyMap[p.ref]?.[m] ?? 0)
        return [p.ref, p.material_text, p.product_range, ...qtys, qtys.reduce((a, b) => a + b, 0)]
      }),
    ]
    const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(';')).join('\n')
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `planning_${year}${search || rangeFilter ? '_filtre' : ''}.csv`
    a.click()
    URL.revokeObjectURL(url)
    onToast('Export CSV téléchargé', 'success')
  }

  // ── Rendu loading ────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div style={{ textAlign: 'center', padding: '60px 0', color: C.textMuted }}>
        <div style={{ width: 22, height: 22, border: `3px solid ${C.blue}`, borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.8s linear infinite', margin: '0 auto 12px' }} />
        Chargement du planning…
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    )
  }

  if (produits.length === 0) {
    return (
      <div style={{ textAlign: 'center', padding: '60px 20px', color: C.textMuted }}>
        <div style={{ fontSize: 36, marginBottom: 12 }}>📋</div>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 6 }}>Planning vide</div>
        <div style={{ fontSize: 13 }}>Les données vont être importées depuis votre fichier Excel.</div>
      </div>
    )
  }

  // ── Rendu principal ──────────────────────────────────────────────────────────
  return (
    <div style={{ paddingBottom: 40 }}>

      {/* ── Modal édition ───────────────────────────────────────────────────── */}
      {editCell && (
        <div
          style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
          onClick={e => { if (e.target === e.currentTarget) setEditCell(null) }}
        >
          <div style={{ background: C.white, borderRadius: 18, padding: 24, width: '100%', maxWidth: 400, boxShadow: C.shadowMd }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 2 }}>{editCell.ref}</div>
            <div style={{ fontSize: 14, fontWeight: 800, color: C.text, marginBottom: 3, lineHeight: 1.3 }}>{editCell.productLabel}</div>
            <div style={{ fontSize: 13, color: C.blue, fontWeight: 600, marginBottom: 20 }}>{moisLabel(editCell.mois)}</div>

            <label style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: 'uppercase', letterSpacing: '0.06em', display: 'block', marginBottom: 6 }}>
              Quantité — actuel&nbsp;: {editCell.oldQte.toLocaleString('fr-FR')}
            </label>
            <input
              ref={editRef}
              type="number"
              min={0}
              value={editQte}
              onChange={e => setEditQte(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter') saveEdit() }}
              style={{ width: '100%', boxSizing: 'border-box', padding: '10px 14px', border: `2px solid ${C.blue}`, borderRadius: 10, fontSize: 22, fontWeight: 800, textAlign: 'center', fontFamily: 'inherit', color: C.text, marginBottom: 14, outline: 'none' }}
            />

            <label style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: 'uppercase', letterSpacing: '0.06em', display: 'block', marginBottom: 6 }}>
              Raison du changement <span style={{ fontWeight: 400, textTransform: 'none' }}>(optionnel)</span>
            </label>
            <textarea
              value={editRaison}
              onChange={e => setEditRaison(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); saveEdit() } }}
              placeholder="Ex : augmentation conso + noel"
              rows={2}
              style={{ width: '100%', boxSizing: 'border-box', padding: '9px 12px', border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 13, fontFamily: 'inherit', resize: 'vertical', color: C.text, marginBottom: 18 }}
            />

            <div style={{ display: 'flex', gap: 8 }}>
              <button
                onClick={() => setEditCell(null)}
                style={{ flex: 1, padding: '10px 0', background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10, cursor: 'pointer', fontSize: 14, fontWeight: 600, color: C.textMuted, fontFamily: 'inherit' }}
              >
                Annuler
              </button>
              <button
                onClick={saveEdit}
                disabled={saving}
                style={{ flex: 2, padding: '10px 0', background: saving ? C.border : C.blue, color: saving ? C.textMuted : '#fff', border: 'none', borderRadius: 10, cursor: saving ? 'default' : 'pointer', fontSize: 14, fontWeight: 700, fontFamily: 'inherit' }}
              >
                {saving ? 'Enregistrement…' : '✓ Enregistrer'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Barre d'outils ──────────────────────────────────────────────────── */}
      <div style={{ marginBottom: 14, display: 'flex', flexDirection: 'column', gap: 8 }}>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <input
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="REF ou nom produit…"
            style={{ flex: 1, minWidth: 150, padding: '9px 12px', border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 13, fontFamily: 'inherit', background: C.white, color: C.text }}
          />
          <select
            value={rangeFilter}
            onChange={e => setRange(e.target.value)}
            style={{ padding: '9px 10px', border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 12, fontFamily: 'inherit', background: C.white, color: C.text }}
          >
            <option value="">Toutes gammes</option>
            {ranges.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
          <select
            value={year}
            onChange={e => setYear(Number(e.target.value))}
            style={{ padding: '9px 10px', border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 12, fontFamily: 'inherit', background: C.white, color: C.text }}
          >
            {[2025, 2026, 2027].map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
          <div style={{ fontSize: 12, color: C.textMuted }}>
            {filtered.length}/{produits.length} produits
            {filtered.length > 200 && <span style={{ color: C.orange }}> · affichage limité à 200</span>}
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            <button
              onClick={() => setView(v => v === 'grid' ? 'historique' : 'grid')}
              style={{ padding: '7px 12px', background: view === 'historique' ? C.purpleSoft : C.bg, border: `1px solid ${view === 'historique' ? C.purple : C.border}`, borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600, color: view === 'historique' ? C.purple : C.textSec, fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 5 }}
            >
              {view === 'historique' ? '← Grille' : '🕒 Historique'}
            </button>
            {view === 'grid' && (
              <button
                onClick={exportCSV}
                style={{ padding: '7px 12px', background: C.greenSoft, border: `1px solid ${C.green}44`, borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600, color: C.green, fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 5 }}
              >
                ⬇ Export CSV
              </button>
            )}
          </div>
        </div>
      </div>

      {/* ── Vue Historique ───────────────────────────────────────────────────── */}
      {view === 'historique' && (
        <div>
          <input
            value={histSearch}
            onChange={e => setHistSearch(e.target.value)}
            placeholder="Filtrer par REF, produit ou raison…"
            style={{ width: '100%', boxSizing: 'border-box', padding: '9px 12px', border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 13, fontFamily: 'inherit', background: C.white, marginBottom: 12 }}
          />
          {histLoading ? (
            <div style={{ textAlign: 'center', padding: 40, color: C.textMuted }}>Chargement…</div>
          ) : filteredHist.length === 0 ? (
            <div style={{ textAlign: 'center', padding: 40, color: C.textMuted }}>Aucune modification enregistrée</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {filteredHist.map(h => {
                const up = h.variation > 0
                const down = h.variation < 0
                return (
                  <div key={h.id} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: '11px 14px', boxShadow: C.shadow }}>
                    <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8 }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap', marginBottom: 3 }}>
                          <span style={{ fontSize: 12, fontWeight: 800, color: C.text, fontFamily: 'monospace' }}>{h.ref}</span>
                          <span style={{ fontSize: 11, color: C.blue, fontWeight: 600, background: C.blueSoft, borderRadius: 6, padding: '1px 6px' }}>{moisLabel(h.mois)}</span>
                        </div>
                        <div style={{ fontSize: 12, color: C.textSec, marginBottom: h.raison ? 3 : 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{h.produit_label}</div>
                        {h.raison && (
                          <div style={{ fontSize: 11, color: C.textMuted, fontStyle: 'italic' }}>"{h.raison}"</div>
                        )}
                      </div>
                      <div style={{ textAlign: 'right', flexShrink: 0 }}>
                        <div style={{ fontSize: 15, fontWeight: 800, color: up ? C.green : down ? C.red : C.textMuted }}>
                          {up ? '+' : ''}{h.variation.toLocaleString('fr-FR')}
                        </div>
                        <div style={{ fontSize: 10, color: C.textMuted, marginTop: 1 }}>
                          {(h.ancienne_qte ?? 0).toLocaleString('fr-FR')} → {h.nouvelle_qte.toLocaleString('fr-FR')}
                        </div>
                      </div>
                    </div>
                    <div style={{ fontSize: 10, color: C.textMuted, marginTop: 7 }}>
                      {new Date(h.modifie_le).toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                      {h.modifie_par && h.modifie_par !== 'Utilisateur' && ` · ${h.modifie_par}`}
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </div>
      )}

      {/* ── Vue Grille ──────────────────────────────────────────────────────── */}
      {view === 'grid' && (
        <div style={{ overflowX: 'auto', borderRadius: 12, border: `1px solid ${C.border}`, background: C.white, boxShadow: C.shadow }}>
          <table style={{ borderCollapse: 'collapse', minWidth: '100%' }}>
            <thead>
              <tr>
                {/* Colonnes fixes */}
                <th style={{ ...thBase, position: 'sticky', left: 0, zIndex: 3, minWidth: 90, maxWidth: 90, textAlign: 'left' }}>REF</th>
                <th style={{ ...thBase, position: 'sticky', left: 90, zIndex: 3, minWidth: 200, textAlign: 'left' }}>Produit</th>
                {/* Colonnes mois */}
                {MONTH_SHORT.map((m, i) => (
                  <th key={i} style={{ ...thBase, minWidth: 58, textAlign: 'center' }}>{m}</th>
                ))}
                {/* Total */}
                <th style={{ ...thBase, minWidth: 68, textAlign: 'center', color: C.teal }}>Total</th>
              </tr>
            </thead>

            <tbody>
              {filtered.slice(0, 200).map((p, idx) => {
                const rowQtys = months.map(m => qtyMap[p.ref]?.[m] ?? 0)
                const rowTotal = rowQtys.reduce((a, b) => a + b, 0)
                const even = idx % 2 === 0
                const rowBg = even ? C.white : '#f9fafb'

                return (
                  <tr key={p.ref} style={{ background: rowBg, borderBottom: `1px solid ${C.border}` }}>
                    {/* REF (sticky) */}
                    <td style={{ ...tdBase, position: 'sticky', left: 0, zIndex: 1, background: rowBg, fontFamily: 'monospace', fontSize: 11, fontWeight: 700, color: C.blue, maxWidth: 90, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {p.ref}
                    </td>
                    {/* Nom (sticky) */}
                    <td style={{ ...tdBase, position: 'sticky', left: 90, zIndex: 1, background: rowBg, maxWidth: 200 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: C.text, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 196 }}>{p.material_text}</div>
                      {p.product_range && <div style={{ fontSize: 10, color: C.textMuted, marginTop: 1 }}>{p.product_range}</div>}
                    </td>
                    {/* Cellules mois */}
                    {months.map((m, i) => {
                      const qty = rowQtys[i]
                      return (
                        <td
                          key={m}
                          title={`${p.ref} · ${moisLabel(m)} · ${qty.toLocaleString('fr-FR')} unités — clic pour modifier`}
                          onClick={() => openEdit(p.ref, m, p.material_text)}
                          style={{
                            ...tdBase,
                            textAlign: 'center',
                            cursor: 'pointer',
                            fontSize: 12,
                            fontWeight: qty > 0 ? 600 : 400,
                            color: qty > 0 ? C.textSec : C.border,
                            minWidth: 58,
                          }}
                          onMouseEnter={e => { e.currentTarget.style.background = C.blueSoft; e.currentTarget.style.color = C.blue }}
                          onMouseLeave={e => { e.currentTarget.style.background = rowBg; e.currentTarget.style.color = qty > 0 ? C.textSec : C.border }}
                        >
                          {fmt(qty)}
                        </td>
                      )
                    })}
                    {/* Total ligne */}
                    <td style={{ ...tdBase, textAlign: 'center', fontSize: 12, fontWeight: 700, color: rowTotal > 0 ? C.teal : C.border }}>
                      {fmt(rowTotal)}
                    </td>
                  </tr>
                )
              })}

              {/* Ligne totaux */}
              <tr style={{ background: C.bg, borderTop: `2px solid ${C.border}` }}>
                <td
                  colSpan={2}
                  style={{ ...tdBase, position: 'sticky', left: 0, zIndex: 1, background: C.bg, fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: 'uppercase', letterSpacing: '0.04em', whiteSpace: 'nowrap' }}
                >
                  TOTAL ({filtered.length} produits)
                </td>
                {monthTotals.map((t, i) => (
                  <td key={i} style={{ ...tdBase, textAlign: 'center', fontSize: 12, fontWeight: 700, color: t > 0 ? C.teal : C.border }}>
                    {fmt(t)}
                  </td>
                ))}
                <td style={{ ...tdBase, textAlign: 'center', fontSize: 13, fontWeight: 800, color: C.teal }}>
                  {grandTotal > 0 ? grandTotal.toLocaleString('fr-FR') : '—'}
                </td>
              </tr>
            </tbody>
          </table>

          {filtered.length > 200 && (
            <div style={{ padding: '8px 14px', fontSize: 11, color: C.orange, borderTop: `1px solid ${C.border}`, fontWeight: 600 }}>
              ⚠ {filtered.length - 200} produits masqués — utilisez la recherche pour affiner
            </div>
          )}
        </div>
      )}

      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  )
}
