-- Filtro do pandoc: renderiza TODAS as tabelas como tabelas "pipe" (|) do GitHub Markdown.
--
-- Sem este filtro o escritor gfm cai para <table> em HTML bruto sempre que a tabela tem
-- larguras relativas, legenda, listas ou células mescladas — ilegível como Markdown.
-- Aqui cada célula vira uma única linha (blocos separados por <br>, listas com "•"),
-- células mescladas (colspan/rowspan) são desfeitas (o texto fica na célula de origem) e a
-- legenda passa a um parágrafo em negrito antes da tabela. Nenhum texto é descartado.

local BR = pandoc.RawInline('html', '<br>')

-- junta os blocos de uma célula numa só lista de inlines
local function achatar(blocks)
  local inlines = pandoc.List()
  local function juntar(novos)
    if #novos == 0 then return end
    if #inlines > 0 then inlines:insert(BR) end
    inlines:extend(novos)
  end
  for _, b in ipairs(blocks) do
    if b.t == 'Plain' or b.t == 'Para' then
      juntar(b.content)
    elseif b.t == 'BulletList' or b.t == 'OrderedList' then
      for _, item in ipairs(b.content) do
        local partes = pandoc.List({ pandoc.Str('•'), pandoc.Space() })
        partes:extend(achatar(item))
        juntar(partes)
      end
    else
      local txt = pandoc.utils.stringify(b)
      if txt ~= '' then juntar(pandoc.List({ pandoc.Str(txt) })) end
    end
  end
  return inlines
end

-- texto gfm de uma célula (formatação inline preservada), sem quebras de linha reais
local function texto_celula(cell)
  local inlines = achatar(cell.contents)
  local s = pandoc.write(pandoc.Pandoc({ pandoc.Plain(inlines) }), 'gfm', { wrap_text = 'none' })
  s = s:gsub('%s*\\?\n%s*', ' '):gsub('^%s+', ''):gsub('%s+$', '')
  return (s:gsub('|', '\\|'))
end

local function linhas_de(rows, destino)
  for _, row in ipairs(rows) do
    local cels = {}
    for _, cell in ipairs(row.cells) do cels[#cels + 1] = texto_celula(cell) end
    destino[#destino + 1] = cels
  end
end

function Table(tbl)
  local cab, corpo = {}, {}
  linhas_de(tbl.head.rows, cab)
  for _, b in ipairs(tbl.bodies) do
    linhas_de(b.head, corpo)
    linhas_de(b.body, corpo)
  end
  linhas_de(tbl.foot.rows, corpo)

  local ncols = #tbl.colspecs
  for _, l in ipairs(cab) do ncols = math.max(ncols, #l) end
  for _, l in ipairs(corpo) do ncols = math.max(ncols, #l) end
  if ncols == 0 then return nil end

  local function fmt(l)
    local c = {}
    for i = 1, ncols do c[i] = l[i] or '' end
    return '| ' .. table.concat(c, ' | ') .. ' |'
  end

  local out = {}
  -- pipe tables exigem cabeçalho: sem ele, usa-se uma linha de cabeçalho vazia
  out[#out + 1] = fmt(cab[1] or {})
  out[#out + 1] = '|' .. string.rep(' --- |', ncols)
  for i = 2, #cab do out[#out + 1] = fmt(cab[i]) end
  for _, l in ipairs(corpo) do out[#out + 1] = fmt(l) end

  local blocos = {}
  local legenda = pandoc.utils.stringify(tbl.caption.long)
  if legenda ~= '' then
    blocos[#blocos + 1] = pandoc.Para({ pandoc.Strong({ pandoc.Str(legenda) }) })
  end
  blocos[#blocos + 1] = pandoc.RawBlock('markdown', table.concat(out, '\n'))
  return blocos
end
