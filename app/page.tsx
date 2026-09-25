'use client'

import { useMemo, useState } from 'react'

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

const initialProducts: Product[] = [
  { id: 1, code: '001', description: 'Café coado', active: true, unit: 'UN', group: 'BEBIDAS', price: '6,50', barcode: '7890000000011', ncm: '0901.21.00', optional: false, prepTime: '05 min' },
  { id: 2, code: '002', description: 'Pão de queijo', active: true, unit: 'UN', group: 'PAES DE QUEIJO', price: '5,00', barcode: '7890000000028', ncm: '1905.90.90', optional: false, prepTime: '08 min' },
  { id: 3, code: '003', description: 'Brigadeiro gourmet', active: true, unit: 'UN', group: 'DOCES', price: '4,50', barcode: '7890000000035', ncm: '1704.90.90', optional: true, prepTime: '—' },
  { id: 4, code: '004', description: 'Sanduíche especial', active: false, unit: 'UN', group: 'ESPECIAL', price: '18,90', barcode: '7890000000042', ncm: '1602.32.00', optional: false, prepTime: '12 min' },
]

const columns = [
  ['code', 'Código'], ['description', 'Descrição'], ['unit', 'Unidade'], ['group', 'Grupo de venda'],
  ['price', 'Valor de venda'], ['barcode', 'Código de barras'], ['ncm', 'NCM'], ['prepTime', 'Tempo preparo'],
] as const

function Icon({ children }: { children: React.ReactNode }) {
  return <span className="icon" aria-hidden="true">{children}</span>
}

export default function Page() {
  const [products, setProducts] = useState(initialProducts)
  const [query, setQuery] = useState('')
  const [activeOnly, setActiveOnly] = useState(false)
  const [selected, setSelected] = useState<number | null>(1)
  const [saved, setSaved] = useState(false)

  const filteredProducts = useMemo(() => products.filter((product) => {
    const matchesQuery = `${product.code} ${product.description} ${product.group}`.toLowerCase().includes(query.toLowerCase())
    return matchesQuery && (!activeOnly || product.active)
  }), [products, query, activeOnly])

  function updateProduct(id: number, key: keyof Product, value: string | boolean) {
    setProducts((current) => current.map((product) => product.id === id ? { ...product, [key]: value } : product))
    setSaved(false)
  }

  function addProduct() {
    const id = Math.max(...products.map((product) => product.id)) + 1
    setProducts((current) => [...current, { id, code: String(id).padStart(3, '0'), description: 'Novo produto', active: true, unit: 'UN', group: 'GERAL', price: '0,00', barcode: '', ncm: '', optional: false, prepTime: '—' }])
    setSelected(id)
    setSaved(false)
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
          <div className="page-heading"><div><div className="eyebrow">CATÁLOGO OPERACIONAL</div><h1>Editor de produtos</h1><p>Edite os produtos no formato padrão da sua empresa.</p></div><div className="heading-actions"><button className="button secondary"><Icon>⇩</Icon> Importar Excel</button><button className="button primary" onClick={addProduct}><Icon>＋</Icon> Novo produto</button></div></div>

          <div className="notice"><div className="notice-icon">i</div><div><strong>Modelo ativo: Produto Venda 1</strong><span>As alterações respeitam as colunas e regras definidas no modelo da empresa.</span></div><button aria-label="Fechar aviso">×</button></div>

          <div className="stats"><div><span className="stat-label">TOTAL DE PRODUTOS</span><strong>{products.length}</strong><small className="positive">↑ 3 este mês</small></div><div><span className="stat-label">ATIVOS</span><strong>{products.filter((product) => product.active).length}</strong><small>disponíveis para venda</small></div><div><span className="stat-label">COM PENDÊNCIAS</span><strong className="warning-text">2</strong><small>campos para revisar</small></div><div><span className="stat-label">MODELO</span><strong className="model-name">Produto Venda 1</strong><small>35 colunas configuradas</small></div></div>

          <section className="editor-card"><div className="card-toolbar"><div className="search-wrap"><Icon>⌕</Icon><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar por código, nome ou grupo..." aria-label="Buscar produtos" /></div><label className="check-label"><input type="checkbox" checked={activeOnly} onChange={(event) => setActiveOnly(event.target.checked)} /> Mostrar apenas ativos</label><button className="filter-button"><Icon>≡</Icon> Filtros <span>2</span></button><button className="view-button" aria-label="Opções de visualização">▦</button></div>
            <div className="table-meta"><span><strong>{filteredProducts.length}</strong> produtos encontrados</span><span className="edit-hint"><span className="blue-dot" /> Clique em qualquer célula para editar</span></div>
            <div className="table-scroll"><table><thead><tr><th className="select-col"><input type="checkbox" aria-label="Selecionar todos" /></th><th>Status</th>{columns.map(([, label]) => <th key={label}>{label}{['Código', 'Descrição', 'Valor de venda'].includes(label) && <span className="sort">↕</span>}</th>)}<th /></tr></thead><tbody>{filteredProducts.map((product) => <tr key={product.id} className={selected === product.id ? 'selected-row' : ''} onClick={() => setSelected(product.id)}><td className="select-col"><input type="checkbox" checked={selected === product.id} onChange={() => setSelected(product.id)} aria-label={`Selecionar ${product.description}`} /></td><td><button className={`toggle ${product.active ? 'on' : ''}`} onClick={(event) => { event.stopPropagation(); updateProduct(product.id, 'active', !product.active) }} aria-label={`Alternar status de ${product.description}`}><span /></button></td>{columns.map(([key]) => <td key={key}><input className="cell-input" value={String(product[key])} onChange={(event) => updateProduct(product.id, key, event.target.value)} onClick={(event) => event.stopPropagation()} aria-label={`${key} de ${product.description}`} /></td>)}<td><button className="row-more" aria-label={`Mais opções para ${product.description}`}>•••</button></td></tr>)}</tbody></table></div>
            <div className="table-footer"><span>Exibindo {filteredProducts.length} de {products.length} produtos</span><div className="pagination"><button disabled>‹</button><button className="current">1</button><button>2</button><button>3</button><button>›</button></div><button className="save-button" onClick={() => setSaved(true)}>{saved ? 'Alterações salvas' : 'Salvar alterações'} <span>⌘ S</span></button></div>
          </section>
          <footer className="page-footer"><span><span className="green-dot" /> Todas as alterações são salvas no histórico</span><span>Modelo v1.4 · Atualizado em 12/06/2024</span></footer>
        </div>
      </section>
    </main>
  )
}
