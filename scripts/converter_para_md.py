"""Converte docx/pptx/pdf/xlsx de uma pasta em Markdown, ao lado de cada original.

Uso:
    python converter_para_md.py <pasta> [--dry-run]

Regras:
- SÃ³ converte ficheiros que ainda nÃ£o tÃªm versÃ£o .md (nunca sobrescreve).
- O .md leva o mesmo nome-base do original. Se dois originais da mesma pasta partilham o
  nome-base (ex.: X.docx e X.pdf), os .md recebem a extensÃ£o no nome: X.docx.md / X.pdf.md.
- SÃ³ texto e tabelas: imagens e formataÃ§Ã£o visual nÃ£o sÃ£o incluÃ­das.
- Suporta caminhos com mais de 260 caracteres (prefixo \\\\?\\ no Windows).

DependÃªncias: pandoc no PATH, python-pptx, openpyxl, PyMuPDF (fitz).
"""
import collections
import datetime
import os
import re
import shutil
import subprocess
import sys
import tempfile

import fitz  # PyMuPDF
import openpyxl
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.stdout.reconfigure(encoding='utf-8')
LP = '\\\\?\\' if os.name == 'nt' else ''
EXTS = {'.docx', '.pptx', '.pdf', '.xlsx'}
FILTRO_TABELAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pandoc_tabelas_simples.lua')
HOJE = datetime.date.today().isoformat()


def cabecalho(nome_original):
    return (f'> Convertido automaticamente de `{nome_original}` em {HOJE}. '
            'Apenas texto e tabelas; imagens e formataÃ§Ã£o visual nÃ£o foram incluÃ­das.\n\n')


def esc_celula(v):
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        v = v.date().isoformat() if v.time() == datetime.time(0) else v.isoformat(sep=' ')
    elif isinstance(v, datetime.date):
        v = v.isoformat()
    s = str(v).strip().replace('\r\n', '\n').replace('\r', '\n')
    return s.replace('|', '\\|').replace('\n', '<br>')


def tabela_md(linhas):
    """linhas: lista de listas de strings; a 1.Âª linha vira o cabeÃ§alho."""
    if not linhas:
        return ''
    n = max(len(l) for l in linhas)
    linhas = [l + [''] * (n - len(l)) for l in linhas]
    out = ['| ' + ' | '.join(linhas[0]) + ' |', '|' + '|'.join([' --- '] * n) + '|']
    out += ['| ' + ' | '.join(l) + ' |' for l in linhas[1:]]
    return '\n'.join(out) + '\n'


# ---------------- docx (pandoc) ----------------
def docx_para_md(caminho):
    tmp = tempfile.mkdtemp(prefix='conv_')
    try:
        # cÃ³pia para um caminho curto: o pandoc/CreateProcess nÃ£o lida bem com >260 caracteres
        curto = os.path.join(tmp, 'in.docx')
        shutil.copyfile(caminho, curto)
        r = subprocess.run(
            ['pandoc', 'in.docx', '-f', 'docx', '-t', 'gfm', '--wrap=none',
             f'--lua-filter={FILTRO_TABELAS}'],
            cwd=tmp, capture_output=True, text=True, encoding='utf-8',
        )
        if r.returncode != 0:
            raise RuntimeError(r.stderr.strip()[:300])
        md = r.stdout
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    # Ã¢ncoras vazias que o pandoc deixa nos tÃ­tulos/Ã­ndices (<span id="_Tocâ€¦" class="anchor"></span>)
    md = re.sub(r'<span [^>]*></span>', '', md)
    md = re.sub(r'</?span[^>]*>', '', md)
    info = {}
    n_html = len(re.findall(r'^<table', md, flags=re.M)) + md.count('[TABLE]')
    if n_html:
        info['tabelas_html'] = n_html
    # imagens: substitui por marcador (com o texto alternativo, se existir)
    md = re.sub(r'!\[([^\]]*)\]\([^)]*\)',
                lambda m: f'*[imagem: {m.group(1).strip()}]*' if m.group(1).strip() else '*[imagem]*', md)
    md = re.sub(r'<img [^>]*?(?:alt="([^"]*)")?[^>]*/?>',
                lambda m: f'*[imagem: {m.group(1)}]*' if m.group(1) else '*[imagem]*', md)
    return md.strip() + '\n', info


# ---------------- pptx ----------------
def _formas(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _formas(sh.shapes)
        else:
            yield sh


def _pos(sh):
    try:
        return (int((sh.top or 0) // 250000), int(sh.left or 0))
    except Exception:
        return (0, 0)


def pptx_para_md(caminho):
    prs = Presentation(caminho)
    out, so_visuais = [], []
    for i, slide in enumerate(prs.slides, 1):
        titulo = None
        try:
            if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
                titulo = slide.shapes.title.text_frame.text.strip() or None
        except Exception:
            pass
        out.append(f'## Slide {i}' + (f': {titulo}' if titulo else '') + '\n')
        corpo, n_imgs = [], 0
        for sh in sorted(_formas(slide.shapes), key=_pos):
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                n_imgs += 1
                continue
            if getattr(sh, 'has_table', False) and sh.has_table:
                linhas = [[esc_celula(c.text) for c in row.cells] for row in sh.table.rows]
                corpo.append(tabela_md(linhas))
                continue
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                corpo.append('*[grÃ¡fico]*\n')
                continue
            if not getattr(sh, 'has_text_frame', False) or not sh.has_text_frame:
                continue
            if titulo and sh == slide.shapes.title:
                continue
            paras = [(p.level, ''.join(r.text for r in p.runs).strip() or p.text.strip())
                     for p in sh.text_frame.paragraphs]
            paras = [(lv, t) for lv, t in paras if t]
            if not paras:
                continue
            if len(paras) == 1:
                corpo.append(paras[0][1] + '\n')
            else:
                corpo.append('\n'.join(f'{"  " * lv}- {t}' for lv, t in paras) + '\n')
        if corpo:
            out.append('\n'.join(corpo))
        elif not titulo:
            out.append('*[slide sem texto â€” conteÃºdo apenas visual]*\n')
            so_visuais.append(i)
        if n_imgs:
            out.append(f'*[{n_imgs} imagem(ns) no slide]*\n')
        if slide.has_notes_slide:
            nota = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ''
            if nota:
                out.append('> **Notas do orador:** ' + nota.replace('\n', '\n> ') + '\n')
    return '\n'.join(out).strip() + '\n', {'slides_so_visuais': so_visuais, 'slides': len(prs.slides)}


# ---------------- pdf ----------------
def pdf_para_md(caminho):
    doc = fitz.open(caminho)
    partes, sem_texto = [], []
    for i, pg in enumerate(doc, 1):
        txt = re.sub(r'[ \t]+\n', '\n', pg.get_text('text', sort=True))
        txt = re.sub(r'\n{3,}', '\n\n', txt).strip()
        if not txt:
            sem_texto.append(i)
            partes.append(f'<!-- pÃ¡gina {i} -->\n*[pÃ¡gina sem texto â€” conteÃºdo apenas visual]*\n')
        else:
            partes.append(f'<!-- pÃ¡gina {i} -->\n{txt}\n')
    return '\n'.join(partes).strip() + '\n', {'paginas': len(doc), 'paginas_sem_texto': sem_texto}


# ---------------- xlsx ----------------
def xlsx_para_md(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    out, vazias = [], []
    for ws in wb.worksheets:
        oculta = ws.sheet_state != 'visible'
        out.append(f'## {ws.title}' + (' (aba oculta)' if oculta else '') + '\n')
        linhas = [[esc_celula(c) for c in row] for row in ws.iter_rows(values_only=True)]
        linhas = [l for l in linhas if any(l)]
        if not linhas:
            out.append('*[aba vazia]*\n')
            vazias.append(ws.title)
            continue
        ncol = max(len(l) for l in linhas)
        keep = [j for j in range(ncol) if any(j < len(l) and l[j] for l in linhas)]
        linhas = [[l[j] if j < len(l) else '' for j in keep] for l in linhas]
        if len(keep) > 12 and len(linhas) <= len(keep) // 3:
            linhas = [list(col) for col in zip(*linhas)]  # transpÃµe tabelas muito largas
            out.append('*(tabela transposta para legibilidade)*\n')
        out.append(tabela_md(linhas))
    return '\n'.join(out).strip() + '\n', {'abas': len(wb.worksheets), 'abas_vazias': vazias}


CONVERSORES = {'.docx': docx_para_md, '.pptx': pptx_para_md, '.pdf': pdf_para_md, '.xlsx': xlsx_para_md}


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    raiz = os.path.abspath(sys.argv[1])
    dry = '--dry-run' in sys.argv
    todos = []
    for r, _, fs in os.walk(LP + raiz):
        for f in sorted(fs):
            todos.append((r, f))
    por_base = collections.defaultdict(list)
    for r, f in todos:
        base, ext = os.path.splitext(f)
        if ext.lower() in EXTS:
            por_base[(r, base)].append(ext.lower())

    contagem = collections.Counter()
    alertas = []
    for r, f in todos:
        base, ext = os.path.splitext(f)
        ext = ext.lower()
        if ext not in EXTS:
            continue
        colide = len(por_base[(r, base)]) > 1
        destino_nome = f'{base}{ext}.md' if colide else f'{base}.md'
        destino = os.path.join(r, destino_nome)
        rotulo = os.path.relpath(os.path.join(r, f), LP + raiz)
        if os.path.exists(destino) or (not colide and os.path.exists(os.path.join(r, base + '.md'))):
            contagem['ja_tinha_md'] += 1
            continue
        if dry:
            print('converteria:', rotulo, '->', destino_nome)
            contagem['a_converter'] += 1
            continue
        try:
            corpo, info = CONVERSORES[ext](os.path.join(r, f))
        except Exception as e:  # noqa: BLE001 - relatamos e seguimos com os outros
            contagem['falhou'] += 1
            alertas.append(f'FALHOU {rotulo}: {e}')
            print('âœ—', rotulo, '-', e)
            continue
        with open(destino, 'w', encoding='utf-8', newline='\n') as fh:
            fh.write(cabecalho(f) + corpo)
        contagem[ext] += 1
        if info.get('slides_so_visuais'):
            alertas.append(f'{rotulo}: {len(info["slides_so_visuais"])}/{info["slides"]} slides sÃ³ com imagem {info["slides_so_visuais"]}')
        if info.get('paginas_sem_texto'):
            alertas.append(f'{rotulo}: {len(info["paginas_sem_texto"])}/{info["paginas"]} pÃ¡ginas sem texto')
        if info.get('tabelas_html'):
            alertas.append(f'{rotulo}: {info["tabelas_html"]} tabela(s) nÃ£o convertida(s) para |')
        if not corpo.strip():
            alertas.append(f'{rotulo}: resultado vazio')
        print('âœ“', rotulo, '->', destino_nome)

    print('\nResumo:', dict(contagem))
    for a in alertas:
        print('ALERTA:', a)


if __name__ == '__main__':
    main()

