// Client Supabase léger — REST API sans dépendance externe
const BASE = 'https://fcjtntvuuhmrqgafdsjl.supabase.co/rest/v1'
const KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZjanRudHZ1dWhtcnFnYWZkc2psIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ1MTI2OTYsImV4cCI6MjA5MDA4ODY5Nn0.dx8b_rkv7Lt-9K-xGq9-z9OnLsolFNnWJfoTTA8re7M'

const H: Record<string, string> = {
  apikey: KEY,
  Authorization: `Bearer ${KEY}`,
  'Content-Type': 'application/json',
}

function buildUrl(table: string, filters: string[], order: string | null, select: string) {
  let url = `${BASE}/${table}?select=${select}`
  for (const f of filters) url += `&${f}`
  if (order) url += `&order=${order}`
  return url
}

// ─── Mutation builder (update/delete + eq chaining) ─────────────────────────

class MutationBuilder {
  private table: string
  private body: object
  private method: 'PATCH' | 'DELETE'
  private filters: string[] = []

  constructor(table: string, method: 'PATCH' | 'DELETE', body: object = {}) {
    this.table = table
    this.method = method
    this.body = body
  }

  eq(col: string, val: unknown): this {
    this.filters.push(`${col}=eq.${val}`)
    return this
  }

  then(resolve: (v: { data: any; error: any }) => void) {
    let url = `${BASE}/${this.table}`
    if (this.filters.length) url += '?' + this.filters.join('&')
    fetch(url, {
      method: this.method,
      headers: { ...H, Prefer: 'return=minimal' },
      body: this.method === 'PATCH' ? JSON.stringify(this.body) : undefined,
    })
      .then(async res => {
        if (!res.ok) { const t = await res.text(); resolve({ data: null, error: { message: t } }) }
        else resolve({ data: null, error: null })
      })
      .catch((e: any) => resolve({ data: null, error: { message: e.message } }))
  }
}

// ─── Select builder ──────────────────────────────────────────────────────────

class SelectBuilder {
  private table: string
  private _select = '*'
  private _filters: string[] = []
  private _order: string | null = null

  constructor(table: string) { this.table = table }

  select(cols: string) { this._select = cols; return this }

  order(col: string, opts?: { ascending?: boolean }) {
    this._order = `${col}.${opts?.ascending === false ? 'desc' : 'asc'}`
    return this
  }

  eq(col: string, val: unknown) { this._filters.push(`${col}=eq.${val}`); return this }

  then(resolve: (v: { data: any; error: any }) => void) {
    fetch(buildUrl(this.table, this._filters, this._order, this._select), { headers: H })
      .then(async res => {
        if (!res.ok) { const t = await res.text(); return resolve({ data: null, error: { message: t } }) }
        resolve({ data: await res.json(), error: null })
      })
      .catch((e: any) => resolve({ data: null, error: { message: e.message } }))
  }

  // upsert / update on the same table object
  upsert(body: object | object[], _opts?: { onConflict?: string }): Promise<{ error: any }> {
    return fetch(`${BASE}/${this.table}`, {
      method: 'POST',
      headers: { ...H, Prefer: 'resolution=merge-duplicates,return=minimal' },
      body: JSON.stringify(body),
    })
      .then(async res => {
        if (!res.ok) { const t = await res.text(); return { error: { message: t } } }
        return { error: null }
      })
      .catch((e: any) => ({ error: { message: e.message } }))
  }

  update(body: object): MutationBuilder {
    return new MutationBuilder(this.table, 'PATCH', body)
  }

  delete(): MutationBuilder {
    return new MutationBuilder(this.table, 'DELETE')
  }
}

// ─── Public API ──────────────────────────────────────────────────────────────

export const supabase = {
  from: (table: string) => new SelectBuilder(table),
}
