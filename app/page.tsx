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
  'Código', 'Descrição', 'Ativo S/N', 'Descrição Detalhada', 'Unidade', 'Grupo Venda', 'Valor Venda',
  'Codigo Classificação', 'Código Produto Principal', 'Acompanhamento?', 'Codigo Linha Produto', 'Opcional?',
  'Codigo Barras', 'NCM', 'LST', 'CEST', 'Código Exceção Tabela IPI', ' Valor Produto Nível Superior?',
  'Codigo Integração', 'Codigo Produto Referencia', 'Mix de produto?', 'Valor Mix', 'Tempo Preparo',
  'Quantidade Fracionada?', 'Considera TC por Produto?', 'Taxa Serviço?', 'Local Consumo AA One',
  'Descricao AA One', 'Maior Dezoito AA One?', 'Tela Oferta AA One?', 'Produto Combinado PDV One?',
  'Exibir produto na pré conta (venda mesa)', 'Produto auxiliar para lançamento', 'Codigo de Integração de Unidade',
  'Cód. Etiqueta QRCode',
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
  const fileInputRef = useRef<HTMLInputElement>(null)

  const filteredProducts = useMemo(() => products.filter((product) => {
    const matchesQuery = `${product.code} ${product.description} ${product.group}`.toLowerCase().includes(query.toLowerCase())
    return matchesQuery && (!activeOnly || product.active)
  }), [products, query, activeOnly])

  function updateProduct(id: number, key: keyof Product, value: string | boolean) {
    setProducts((current) => current.map((product) => product.id === id ? { ...product, [key]: value } : product))
    setSaved(false)
  }

  async function handleExcelUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return

    const workbook = read(await file.arrayBuffer(), { cellDates: true })
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]
    const rows = utils.sheet_to_json<SpreadsheetRow>(worksheet, { defval: '' })
    const matrix = utils.sheet_to_json<(string | number | boolean | null)[]>(worksheet, { header: 1, defval: '' })
    const headers = rows.length > 0 ? Object.keys(rows[0]) : ((matrix[0] ?? []).map(String).filter(Boolean))
    const importedColumns = headers.map((header) => ({ key: header, label: header }))

    if (headers.length > 0) {
      setSheetColumns(importedColumns)
    }

    if (rows.length > 0 && headers.length > 0) {
      setSheetColumns(importedColumns)
      setProducts(rows.map((row, index) => ({
        id: index + 1,
        code: String(row[headers[0]] ?? '').trim(),
        description: String(row[headers[1]] ?? `Produto ${index + 1}`).trim(),
        active: true,
        unit: String(row[headers[2]] ?? 'UN').trim(),
        group: String(row[headers[3]] ?? 'GERAL').trim(),
        price: String(row[headers[4]] ?? '').trim(),
        barcode: String(row[headers[5]] ?? '').trim(),
        ncm: String(row[headers[6]] ?? '').trim(),
        optional: false,
        prepTime: String(row[headers[7]] ?? '—').trim(),
        ...row,
      })))
      setModelName(file.name.replace(/\\.xlsx?$/i, ''))
      setSelected(rows.length > 0 ? 1 : null)
      setSaved(false)
    }

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
      sheetColumns.map((column) => [column.label, product[column.key as keyof Product] ?? '']),
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
          <a className="nav-item" href="#"><Icon>⌂</Icon> Visão geral</a>
          <a className="nav-item active" href="#editor"><Icon>▦</Icon> Produtos e vendas <b>24</b></a>
          <a className="nav-item" href="#"><Icon>⇄</Icon> Importações</a>
          <a className="nav-item" href="#"><Icon>↗</Icon> Exportações</a>
          <p className="nav-label space">CONFIGURAÇÕES</p>
          <a className="nav-item" href="#"><Icon>⚙</Icon> Regras do cadastro</a>
          <a className="nav-item" href="#"><Icon>?</Icon> Central de ajuda</a>
        </nav>
        <div className="sidebar-bottom"><div className="avatar">MS</div><div><strong>Marina Silva</strong><small>Administradora</small></div><button className="more" aria-label="Mais opções">•••</button></div>
      </aside>

      <section className="workspace" id="editor">
        <header className="topbar"><div className="breadcrumb"><span>Produtos e vendas</span><i>/</i><strong>Editor de produtos</strong></div><div className="top-actions"><span className="status-dot" /> Última sincronização há 2 min <button className="help" aria-label="Ajuda">?</button></div></header>
        <div className="content">
          <div className="page-heading"><div><div className="eyebrow">CATÁLOGO OPERACIONAL</div><h1>Editor de produtos</h1><p>Edite os produtos no formato padrão da sua empresa.</p></div><div className="heading-actions"><input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} hidden /><button className="button secondary" onClick={() => fileInputRef.current?.click()}><Icon>⇩</Icon> Importar Excel</button><button className="button secondary" onClick={exportExcel}><Icon>⇧</Icon> Exportar Excel</button><button className="button primary" onClick={addProduct}><Icon>＋</Icon> Novo produto</button></div></div>

          <div className="notice"><div className="notice-icon">i</div><div><strong>Modelo ativo: {modelName}</strong><span>As alterações respeitam as colunas e regras definidas no modelo da empresa.</span></div><button aria-label="Fechar aviso">×</button></div>

          <div className="stats"><div><span className="stat-label">TOTAL DE PRODUTOS</span><strong>{products.length}</strong><small className="positive">↑ 3 este mês</small></div><div><span className="stat-label">ATIVOS</span><strong>{products.filter((product) => product.active).length}</strong><small>disponíveis para venda</small></div><div><span className="stat-label">COM PENDÊNCIAS</span><strong className="warning-text">2</strong><small>campos para revisar</small></div><div><span className="stat-label">MODELO</span><strong className="model-name">{modelName}</strong><small>35 colunas configuradas</small></div></div>

          <section className="editor-card"><div className="card-toolbar"><div className="search-wrap"><Icon>⌕</Icon><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar por código, nome ou grupo..." aria-label="Buscar produtos" /></div><label className="check-label"><input type="checkbox" checked={activeOnly} onChange={(event) => setActiveOnly(event.target.checked)} /> Mostrar apenas ativos</label><button className="filter-button"><Icon>≡</Icon> Filtros <span>2</span></button><button className="view-button" aria-label="Opções de visualização">▦</button></div>
            <div className="table-meta"><span><strong>{filteredProducts.length}</strong> produtos encontrados</span><span className="edit-hint"><span className="blue-dot" /> Clique em qualquer célula para editar</span></div>
            <div className="table-scroll"><table><thead><tr><th className="select-col"><input type="checkbox" aria-label="Selecionar todos" /></th><th>Status</th>{sheetColumns.map((column) => <th key={column.key}>{column.label}{['Código', 'Descrição', 'Valor de venda'].includes(column.label) && <span className="sort">↕</span>}</th>)}<th /></tr></thead><tbody>{filteredProducts.map((product) => <tr key={product.id} className={selected === product.id ? 'selected-row' : ''} onClick={() => setSelected(product.id)}><td className="select-col"><input type="checkbox" checked={selected === product.id} onChange={() => setSelected(product.id)} aria-label={`Selecionar ${product.description}`} /></td><td><button className={`toggle ${product.active ? 'on' : ''}`} onClick={(event) => { event.stopPropagation(); updateProduct(product.id, 'active', !product.active) }} aria-label={`Alternar status de ${product.description}`}><span /></button></td>{sheetColumns.map((column) => <td key={column.key}><input className="cell-input" value={String(product[column.key as keyof Product] ?? '')} onChange={(event) => updateProduct(product.id, column.key as keyof Product, event.target.value)} onClick={(event) => event.stopPropagation()} aria-label={`${column.label} de ${product.description}`} /></td>)}<td><button className="row-more" aria-label={`Mais opções para ${product.description}`}>•••</button></td></tr>)}</tbody></table></div>
            <div className="table-footer"><span>Exibindo {filteredProducts.length} de {products.length} produtos</span><div className="pagination"><button disabled>‹</button><button className="current">1</button><button>2</button><button>3</button><button>›</button></div><button className="save-button" onClick={() => setSaved(true)}>{saved ? 'Alterações salvas' : 'Salvar alterações'} <span>⌘ S</span></button></div>
          </section>
          <footer className="page-footer"><span><span className="green-dot" /> Todas as alterações são salvas no histórico</span><span>Modelo v1.4 · Atualizado em 12/06/2024</span></footer>
        </div>
      </section>
    </main>
  )
}
