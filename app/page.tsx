'use client'

import { useMemo, useRef, useState } from 'react'
import { read, utils, writeFile } from 'xlsx'

 type SpreadsheetRow = Record<string, string | number | boolean>
 type Column = { key: string; label: string }

type Product = {
  id: number
  code: string
  description: string
  active: boolean
  unit: string
  group: string
  price: string
  barcode: string
  ncm: string
  optional: boolean
  prepTime: string
}

const initialProducts: Product[] = []

const templateHeaders = [
  'Código', 'Descrição do Item', 'Ativo S/N', 'Despesa', 'Tipo de Item', 'Grupo de Item',
  'Genero do Produto', 'Unidade de medida', 'Compõe produto de venda S/N', 'Manufaturado S/N',
  'Valor Unitario', 'Codigo de Barras', 'Codigos de Barras Auxiliares (separados por vírgula)', 'Ncm',
  'Código Serviço (LST)', 'Código .exc. Tabela IPI', 'Codigo Integracao Item', 'Cest', 'Nve',
  'Imposto Federal %', 'Imposto Estadual %', 'Imposto Municipal %', 'Layout Tabela (IBPT)',
  'Código de int. deposito', 'Codigo de Integracao 2', 'Codigo de Integracao Unidade',
]

const defaultColumns: Column[] = templateHeaders.map((header) => ({ key: header, label: header }))

function Icon({ children }: { children: React.ReactNode }) {
  return <span className="icon" aria-hidden="true">{children}</span>
}

export default function Page() {
  const [products, setProducts] = useState(initialProducts)
  const [query, setQuery] = useState('')
  const [activeOnly, setActiveOnly] = useState(false)
  const [selected, setSelected] = useState<number | null>(null)
  const [saved, setSaved] = useState(false)
  const [modelName, setModelName] = useState('Modelo vazio')
  const [sheetColumns, setSheetColumns] = useState<Column[]>(defaultColumns)
  const [sourceName, setSourceName] = useState('Nenhuma base importada')
  const [activeView, setActiveView] = useState<'overview' | 'editor' | 'imports' | 'help'>('overview')
  const [importHistory, setImportHistory] = useState<{ name: string; rows: number; date: string }[]>([])
  const fileInputRef = useRef<HTMLInputElement>(null)
  const sourceInputRef = useRef<HTMLInputElement>(null)

  const filteredProducts = useMemo(() => products.filter((product) => {
    const matchesQuery = `${product.code} ${product.description} ${product.group}`.toLowerCase().includes(query.toLowerCase())
    return matchesQuery && (!activeOnly || product.active)
  }), [products, query, activeOnly])

  function updateProduct(id: number, key: keyof Product, value: string | boolean) {
    const nextValue = typeof value === 'string' && (/^ncm$/i.test(String(key)) || /^cest$/i.test(String(key)))
      ? normalizeFiscalCode(value)
      : value
    setProducts((current) => current.map((product) => product.id === id ? { ...product, [key]: nextValue } : product))
    setSaved(false)
  }

  function cleanValue(value: unknown) {
    return String(value ?? '')
      .replace(/^'+|'+$/g, '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase()
  }

  function cleanBarcode(value: unknown) {
    return String(value ?? '').replace(/^'+|'+$/g, '').replace(/\D/g, '')
  }

  function cleanNumeric(value: unknown) {
    return String(value ?? '')
      .replace(/^'+|'+$/g, '')
      .replace(/[^0-9,.-]/g, '')
  }

  function normalizeFiscalCode(value: unknown, length = 8) {
    const cleaned = cleanBarcode(value)
    return cleaned ? cleaned.padStart(length, '0').slice(-length) : ''
  }

  function normalizeTemplateValue(column: string, value: unknown) {
    if (/^ncm$/i.test(column)) return normalizeFiscalCode(value, 8)
    if (/^cest$/i.test(column)) return normalizeFiscalCode(value, 7)
    if (/c[oó]digo de barras/i.test(column)) return cleanBarcode(value)
    if (/valor unit[aá]rio|valor de venda|pre[cç]o/i.test(column)) return cleanNumeric(value)
    return cleanValue(value)
  }

  function buildProduct(row: SpreadsheetRow, index: number): Product {
    const code = cleanValue(row['Código'] ?? row['cod_interno'])
    const description = cleanValue(row['Descrição'] ?? row['Descrição do Item'] ?? row.desc)
    return {
      id: index + 1,
      code,
      description,
      active: cleanValue(row['Ativo S/N']) !== 'N',
      unit: cleanValue(row['Unidade'] ?? row['Unidade de medida'] ?? row.unidade),
      group: cleanValue(row['Grupo Venda'] ?? row['Grupo de Item'] ?? row.FAMILIA),
      price: cleanNumeric(row['Valor Venda'] ?? row['Valor Unitario'] ?? row['preço de custo']),
      barcode: cleanBarcode(row['Codigo Barras'] ?? row['Codigo de Barras'] ?? row.codbarra),
      ncm: normalizeFiscalCode(row.NCM ?? row.Ncm, 8),
      optional: cleanValue(row['Opcional?']) === 'S',
      prepTime: cleanValue(row['Tempo Preparo']),
      ...row,
    }
  }

  async function handleTemplateUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    const workbook = read(await file.arrayBuffer(), { cellDates: true })
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]
    const matrix = utils.sheet_to_json<(string | number | boolean | null)[]>(worksheet, { header: 1, defval: '' })
    const headers = (matrix[0] ?? []).map(String).filter(Boolean)
    if (headers.length) setSheetColumns(headers.map((header) => ({ key: header, label: header })))
    setProducts([])
    setModelName(file.name.replace(/\\.xlsx?$/i, ''))
    setSourceName('Nenhuma base importada')
    setSelected(null)
    setSaved(false)
    event.target.value = ''
  }

  async function handleSourceUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    const workbook = read(await file.arrayBuffer(), { cellDates: true })
    const sheetName = workbook.SheetNames.find((name) => name.toLowerCase().includes('mater'))
    if (!sheetName) return
    const rows = utils.sheet_to_json<SpreadsheetRow>(workbook.Sheets[sheetName], { defval: '' })
    const mappedRows = rows.map((row, index) => {
      const mapped: SpreadsheetRow = {
        'Código': cleanValue(row.cod_interno),
        'Descrição do Item': cleanValue(row.desc),
        'Ativo S/N': 'S',
        'Tipo de Item': '1',
        'Unidade de medida': cleanValue(row.unidade),
        'Valor Unitario': cleanNumeric(row['preço de custo']),
        'Codigo de Barras': cleanBarcode(row.codbarra),
        'Ncm': normalizeFiscalCode(row.NCM, 8),
        'Cest': normalizeFiscalCode(row.CEST, 7),
      }
      return { ...mapped, id: index + 1 }
    })
    setProducts(mappedRows.map(buildProduct))
    setSourceName(`${rows.length} matérias-primas importadas`)
    setImportHistory((current) => [{ name: file.name, rows: rows.length, date: new Date().toLocaleString('pt-BR') }, ...current])
    setSelected(rows.length ? 1 : null)
    setSaved(false)
    event.target.value = ''
  }

  function addProduct() {
    const id = products.length > 0 ? Math.max(...products.map((product) => product.id)) + 1 : 1
    setProducts((current) => [...current, { id, code: String(id).padStart(3, '0'), description: 'Novo produto', active: true, unit: 'UN', group: 'GERAL', price: '0,00', barcode: '', ncm: '', optional: false, prepTime: '—' }])
    setSelected(id)
    setSaved(false)
  }

  function exportExcel() {
    const rows = products.map((product) => Object.fromEntries(
      sheetColumns.map((column) => [column.label, normalizeTemplateValue(column.label, product[column.key as keyof Product] ?? '')]),
    ))
    const worksheet = utils.json_to_sheet(rows, { header: sheetColumns.map((column) => column.label) })
    const workbook = utils.book_new()
    utils.book_append_sheet(workbook, worksheet, 'Produtos')
    writeFile(workbook, `${modelName || 'modelo-produtos'}-editado.xlsx`)
    setSaved(true)
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand"><div className="brand-mark">C</div><div><strong>cadastro<span>+</span></strong><small>Gestão operacional</small></div></div>
        <nav aria-label="Navegação principal">
          <p className="nav-label">MENU PRINCIPAL</p>
          <button className={`nav-item ${activeView === 'overview' ? 'active' : ''}`} onClick={() => setActiveView('overview')}><Icon>⌂</Icon> Visão geral</button>
          <button className={`nav-item ${activeView === 'editor' ? 'active' : ''}`} onClick={() => setActiveView('editor')}><Icon>▦</Icon> Produtos e vendas <b>{products.length}</b></button>
          <button className={`nav-item ${activeView === 'imports' ? 'active' : ''}`} onClick={() => setActiveView('imports')}><Icon>⇄</Icon> Importações</button>
          <button className="nav-item" onClick={() => setActiveView('editor')}><Icon>↗</Icon> Exportações</button>
          <p className="nav-label space">SUPORTE</p>
          <button className={`nav-item ${activeView === 'help' ? 'active' : ''}`} onClick={() => setActiveView('help')}><Icon>?</Icon> Central de ajuda</button>
        </nav>
        <div className="sidebar-bottom"><div className="avatar">MS</div><div><strong>Marina Silva</strong><small>Administradora</small></div><button className="more" aria-label="Mais opções">•••</button></div>
      </aside>

      <section className="workspace" id="editor">
        <header className="topbar"><div className="breadcrumb"><span>Produtos e vendas</span><i>/</i><strong>Editor de produtos</strong></div><div className="top-actions"><span className="status-dot" /> Última sincronização há 2 min <button className="help" aria-label="Ajuda">?</button></div></header>
        <div className="content">
          {activeView === 'overview' && <section className="overview-panel"><div className="eyebrow">PAINEL DE CONTROLE</div><h1>Visão geral</h1><p>Acompanhe sua operação de cadastro em um só lugar.</p><div className="overview-grid"><div className="overview-card"><span>MATÉRIAS-PRIMAS</span><strong>{products.length}</strong><small>itens no modelo atual</small></div><div className="overview-card"><span>IMPORTAÇÕES</span><strong>{importHistory.length}</strong><small>arquivos processados</small></div><div className="overview-card"><span>MODELO ATIVO</span><strong className="overview-model">{modelName}</strong><small>estrutura configurada</small></div></div><div className="overview-actions"><button className="button primary" onClick={() => setActiveView('editor')}>Abrir editor de produtos</button><button className="button secondary" onClick={() => setActiveView('imports')}>Ver histórico de importações</button></div></section>}
          {activeView === 'imports' && <section className="overview-panel"><div className="eyebrow">RASTREABILIDADE</div><h1>Histórico de importações</h1><p>Consulte os arquivos utilizados para preencher o modelo.</p><div className="history-card">{importHistory.length === 0 ? <div className="empty-state"><strong>Nenhuma importação realizada</strong><span>Quando você enviar uma planilha de matérias-primas, ela aparecerá aqui.</span></div> : importHistory.map((item) => <div className="history-row" key={`${item.name}-${item.date}`}><div className="file-badge">XLSX</div><div><strong>{item.name}</strong><small>{item.rows} matérias-primas · {item.date}</small></div><span className="history-status">Concluída</span></div>)}</div></section>}
          {activeView === 'help' && <section className="overview-panel help-panel"><div className="eyebrow">CENTRAL DE AJUDA</div><h1>Como usar a plataforma</h1><p>Veja o passo a passo para configurar e exportar seu cadastro.</p><div className="help-grid"><div className="tutorial-video" aria-label="Vídeo demonstrativo"><div className="video-screen"><div className="play-circle">▶</div><span>Vídeo demonstrativo</span><small>Importe, revise e exporte em poucos passos</small></div><div className="video-controls"><span>0:42</span><div className="video-line"/><span>1×</span></div></div><div className="tutorial-copy"><h2>Passo a passo</h2><ol><li><strong>Importe o modelo vazio</strong><span>Envie a planilha que define as colunas da sua empresa.</span></li><li><strong>Envie as matérias-primas</strong><span>O sistema encaixa somente os dados nos campos existentes.</span></li><li><strong>Revise e exporte</strong><span>Confira os dados tratados e clique em Exportar Excel.</span></li></ol></div></div></section>}
          <div className={`editor-only ${activeView === 'editor' ? '' : 'hidden-view'}`}><div className="page-heading"><div><div className="eyebrow">CATÁLOGO OPERACIONAL</div><h1>Editor de matérias-primas</h1><p>Use o modelo vazio como base e preencha somente os campos existentes.</p></div><div className="heading-actions"><input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleTemplateUpload} hidden /><input ref={sourceInputRef} type="file" accept=".xlsx,.xls" onChange={handleSourceUpload} hidden /><button className="button secondary" onClick={() => fileInputRef.current?.click()}><Icon>⇩</Icon> Importar modelo vazio</button><button className="button secondary" onClick={() => sourceInputRef.current?.click()}><Icon>⇩</Icon> Importar matérias-primas</button><button className="button secondary" onClick={exportExcel}><Icon>⇧</Icon> Exportar Excel</button></div></div>

          <div className="notice"><div className="notice-icon">i</div><div><strong>Modelo ativo: {modelName}</strong><span>As alterações respeitam as colunas e regras definidas no modelo da empresa.</span></div><button aria-label="Fechar aviso">×</button></div>

          <div className="stats"><div><span className="stat-label">TOTAL DE PRODUTOS</span><strong>{products.length}</strong><small className="positive">↑ 3 este mês</small></div><div><span className="stat-label">ATIVOS</span><strong>{products.filter((product) => product.active).length}</strong><small>disponíveis para venda</small></div><div><span className="stat-label">COM PENDÊNCIAS</span><strong className="warning-text">2</strong><small>campos para revisar</small></div><div><span className="stat-label">MODELO</span><strong className="model-name">{modelName}</strong><small>35 colunas configuradas</small></div></div>

          <section className="editor-card"><div className="card-toolbar"><div className="search-wrap"><Icon>⌕</Icon><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar por código, nome ou grupo..." aria-label="Buscar produtos" /></div><label className="check-label"><input type="checkbox" checked={activeOnly} onChange={(event) => setActiveOnly(event.target.checked)} /> Mostrar apenas ativos</label><button className="filter-button"><Icon>≡</Icon> Filtros <span>2</span></button><button className="view-button" aria-label="Opções de visualização">▦</button></div>
            <div className="table-meta"><span><strong>{filteredProducts.length}</strong> produtos encontrados</span><span className="edit-hint"><span className="blue-dot" /> Clique em qualquer célula para editar</span></div>
            <div className="table-scroll"><table><thead><tr><th className="select-col"><input type="checkbox" aria-label="Selecionar todos" /></th><th>Status</th>{sheetColumns.map((column) => <th key={column.key}>{column.label}{['Código', 'Descrição', 'Valor de venda'].includes(column.label) && <span className="sort">↕</span>}</th>)}<th /></tr></thead><tbody>{filteredProducts.map((product) => <tr key={product.id} className={selected === product.id ? 'selected-row' : ''} onClick={() => setSelected(product.id)}><td className="select-col"><input type="checkbox" checked={selected === product.id} onChange={() => setSelected(product.id)} aria-label={`Selecionar ${product.description}`} /></td><td><button className={`toggle ${product.active ? 'on' : ''}`} onClick={(event) => { event.stopPropagation(); updateProduct(product.id, 'active', !product.active) }} aria-label={`Alternar status de ${product.description}`}><span /></button></td>{sheetColumns.map((column) => <td key={column.key}><input className="cell-input" value={String(product[column.key as keyof Product] ?? '')} onChange={(event) => updateProduct(product.id, column.key as keyof Product, event.target.value)} onClick={(event) => event.stopPropagation()} aria-label={`${column.label} de ${product.description}`} /></td>)}<td><button className="row-more" aria-label={`Mais opções para ${product.description}`}>•••</button></td></tr>)}</tbody></table></div>
            <div className="table-footer"><span>Exibindo {filteredProducts.length} de {products.length} produtos</span><div className="pagination"><button disabled>‹</button><button className="current">1</button><button>2</button><button>3</button><button>›</button></div><button className="save-button" onClick={() => setSaved(true)}>{saved ? 'Alterações salvas' : 'Salvar alterações'} <span>⌘ S</span></button></div>
          </section>
          <footer className="page-footer"><span><span className="green-dot" /> Todas as alterações são salvas no histórico</span><span>Modelo v1.4 · Atualizado em 12/06/2024</span></footer></div>
        </div>
      </section>
    </main>
  )
}
