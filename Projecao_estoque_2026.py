"""
PROJEÇÃO DE ESTOQUE FINANCEIRO - EQTL 2026
===========================================
Lógica (separada em duas etapas independentes):

  ETAPA 1 — QUANTIDADE (não depende de preço/PU em nada):
    Qtd final = Qtd inicial + Qtd entrada (pedidos) - Qtd consumo (plano)
    Consumo é limitado ao disponível (nunca fica negativo).
    Propaga mês a mês.

  ETAPA 2 — VALORAÇÃO (aplicada só no final, sobre as quantidades já prontas):
    R$ = Quantidade × PU
    PU = custo médio real do estoque (Valor/Qtd) quando existe;
         senão, preço da tabela PU; senão, preço do próprio pedido.

  Separar assim evita que qualquer problema de leitura da tabela de PU
  (aba errada, coluna renomeada, etc.) afete o cálculo de quantidade —
  o pior caso é um material ficar com PU = 0 (valor R$ zerado), mas a
  quantidade projetada nunca é afetada.

  ⚙️  Para atualizar todo mês: altere apenas MES_INICIO abaixo.
"""

from pathlib import Path
import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────
MES_INICIO  = '2026-08-01'   # ← atualizar mensalmente

PASTA_BASE  = Path(r'\\10.7.1.90\Expansao_At_Automacao\9. EXEC_CONTROLE OBRAS AT\BASES\PROJEÇÃO_ESTOQUE')
ARQ_PLANO   = Path(r'\\10.7.1.90\Expansao_At_Automacao\5.EXEC_GESTÃO DE PROJETOS_AT\5 - Aquisição de Materiais\3 - Solicitações de Compra\3 Compra Prévia\ATUAL BI\ATUAL PLANO IRRESTRITO\PLANO IRRESTRITO VERSÃO ATUAL BI.xlsm')
ARQ_ENTRADA = PASTA_BASE / 'entrada.xlsx'

# Nomes de aba são strings, não Path!
ABA_ESTOQUE = 'Estoque'
ABA_PEDIDOS = 'PEDIDOS'
ABA_PLANO   = 'LISTA + RESERVAS'
ABA_PU      = 'PU'

ARQ_SAIDA   = PASTA_BASE / 'PROJECAO_ESTOQUE_FINANCEIRO_2026.xlsx'

SUBCLASSES_EXCLUIR = ['TORRE METÁLICA', 'TRANSFORMADOR DE FORÇA', 'MÓDULO GIS', 'REGULADOR DE TENSÃO','POSTE MONOTUBULAR']
CLASSES_EXCLUIR    = ['PRÉ-MOLDADO','TRANSFORMADOR DE FORÇA']

TODOS_MESES = pd.date_range('2026-01-01', '2026-12-01', freq='MS')
CORTE       = pd.Timestamp(MES_INICIO)
MESES       = [m for m in TODOS_MESES if m >= CORTE]


def mes_ptbr(m):
    return m.strftime('%b/%y').upper()\
        .replace('MAY','MAI').replace('AUG','AGO')\
        .replace('SEP','SET').replace('OCT','OUT')\
        .replace('DEC','DEZ')


def normaliza_cod(valor):
    """Remove .0 de floats, espaços, e garante string limpa comparável entre abas."""
    s = str(valor).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def encontra_coluna(df, candidatos):
    """Procura a primeira coluna existente entre os nomes candidatos
    (ignorando maiúsculas/minúsculas e espaços nas bordas). Retorna
    o nome REAL da coluna no df, ou None se nenhum candidato existir."""
    cols_norm = {c.strip().upper(): c for c in df.columns}
    for cand in candidatos:
        alvo = cand.strip().upper()
        if alvo in cols_norm:
            return cols_norm[alvo]
    return None


def ler_aba_com_cabecalho_flexivel(caminho, aba, candidatos_coluna_chave, max_linhas_teste=6, **kwargs):
    """Lê uma aba testando header=0,1,2,...,max_linhas_teste até achar uma
    leitura em que pelo menos uma das colunas candidatas apareça. Isso evita
    quebrar quando o layout da planilha muda (título mesclado, linha extra, etc.).
    Retorna (df, header_usado, coluna_encontrada) ou lança erro descritivo."""
    for header in range(max_linhas_teste):
        try:
            df_tmp = pd.read_excel(caminho, sheet_name=aba, header=header, **kwargs)
        except Exception:
            continue
        df_tmp.columns = df_tmp.columns.astype(str).str.strip()
        col = encontra_coluna(df_tmp, candidatos_coluna_chave)
        if col is not None:
            return df_tmp, header, col

    # Nenhum header testado funcionou — mostra o que existe no header=0 para diagnóstico
    df_diag = pd.read_excel(caminho, sheet_name=aba, header=0, **kwargs)
    raise KeyError(
        f"Não encontrei nenhuma das colunas {candidatos_coluna_chave} na aba '{aba}' "
        f"testando header de 0 a {max_linhas_teste - 1}. "
        f"Colunas encontradas com header=0: {df_diag.columns.tolist()}"
    )


# ─────────────────────────────────────────────
# LEITURA
# ─────────────────────────────────────────────
def ler_bases():
    print(f"📂 Lendo bases... (projeção: {mes_ptbr(CORTE)} → DEZ/26)")

    # ── ESTOQUE ──────────────────────────────
    df_est = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_ESTOQUE, header=0)
    df_est['cod']       = df_est['CÓD MATERIAL'].apply(normaliza_cod)
    df_est['empresa']   = df_est['EMPRESA'].astype(str).str.strip()
    df_est['chave']     = df_est['cod'] + '|' + df_est['empresa']
    df_est['descricao'] = df_est['DESCRIÇÃO MATERIAL'].astype(str).str.strip() if 'DESCRIÇÃO MATERIAL' in df_est.columns else ''
    df_est['Qtd_estoque']   = pd.to_numeric(df_est['QUANTIDADE'], errors='coerce').fillna(0)
    df_est['Valor_estoque'] = pd.to_numeric(df_est['VALOR'],      errors='coerce').fillna(0)

    df_meta = df_est.groupby('chave', as_index=False).agg(cod=('cod','first'), empresa=('empresa','first'), descricao=('descricao','first'))
    df_est  = df_est.groupby('chave', as_index=False).agg(Qtd_estoque=('Qtd_estoque','sum'), Valor_estoque=('Valor_estoque','sum'))
    df_est  = df_est.merge(df_meta, on='chave', how='left')
    print(f"  ✅ Estoque: {len(df_est)} materiais — R$ {df_est['Valor_estoque'].sum():,.0f}")

    negativos = df_est[df_est['Qtd_estoque'] < 0]
    if not negativos.empty:
        print(f"  ⚠️  Materiais com Qtd_estoque negativa: {len(negativos)} — R$ {negativos['Valor_estoque'].sum():,.0f} (tratados como 0 na projeção)")

    # ── PEDIDOS ───────────────────────────────
    df_ped = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_PEDIDOS, header=0)
    df_ped['cod_mat']      = df_ped['Material'].apply(normaliza_cod)
    df_ped['empresa']      = df_ped['EMPRESA'].astype(str).str.strip()
    df_ped['chave']        = df_ped['cod_mat'] + '|' + df_ped['empresa']
    df_ped['Data_entrega'] = pd.to_datetime(df_ped['Dat.rem.estatística'], dayfirst=True, errors='coerce')
    df_ped['Qtd_pedido']   = pd.to_numeric(df_ped['Qtd.a fornecer'], errors='coerce').fillna(0)
    df_ped['Val_pedido']   = pd.to_numeric(df_ped['Valor'],       errors='coerce').fillna(0)
    df_ped['Mes_entrega']  = df_ped['Data_entrega'].dt.to_period('M').dt.to_timestamp()

    if 'CLASSE' in df_ped.columns:
        excl = df_ped[df_ped['CLASSE'].astype(str).str.strip().isin(CLASSES_EXCLUIR)]
        print(f"  🚫 Pedidos — excluídos por CLASSE: {len(excl)} linhas")
        df_ped = df_ped[~df_ped['CLASSE'].astype(str).str.strip().isin(CLASSES_EXCLUIR)].copy()
    if 'SUBCLASSE' in df_ped.columns:
        excl = df_ped[df_ped['SUBCLASSE'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)]
        print(f"  🚫 Pedidos — excluídos por SUBCLASSE: {len(excl)} linhas")
        df_ped = df_ped[~df_ped['SUBCLASSE'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)].copy()

    df_ped = df_ped[(df_ped['Qtd_pedido'] > 0) & (df_ped['Mes_entrega'] >= CORTE)].copy()

    entradas = df_ped.groupby(['chave', 'Mes_entrega'], as_index=False).agg(
        Qtd_entrada=('Qtd_pedido', 'sum'),
        Val_entrada=('Val_pedido', 'sum')
    )
    print(f"  ✅ Pedidos: {len(entradas)} entradas mensais — R$ {df_ped['Val_pedido'].sum():,.0f}")

    # ── PLANO IRRESTRITO (define a QUANTIDADE de consumo) ──────
    df_plano, header_usado, col_cod_sap = ler_aba_com_cabecalho_flexivel(
        ARQ_PLANO, ABA_PLANO,
        candidatos_coluna_chave=['COD SAP', 'COD', 'CÓD MATERIAL', 'Código', 'Material'],
        engine='openpyxl'
    )
    print(f"  🔎 Plano — cabeçalho encontrado na linha {header_usado} (coluna de código: '{col_cod_sap}')")

    col_empresa = encontra_coluna(df_plano, ['EMPRESA', 'Empresa'])
    col_qtd     = encontra_coluna(df_plano, ['QTD ITEM', 'QTD', 'Quantidade'])
    col_data    = encontra_coluna(df_plano, ['DATA NECESSIDADE', 'Data Necessidade'])
    col_classe  = encontra_coluna(df_plano, ['CLASSE MATERIAL', 'CLASSE'])
    col_subcls  = encontra_coluna(df_plano, ['SUBCLASSE MATERIAL', 'SUBCLASSE'])
    col_blq     = encontra_coluna(df_plano, ['BLQ?', 'BLOQUEADO'])

    faltando = [nome for nome, col in [('EMPRESA', col_empresa), ('QTD ITEM', col_qtd), ('DATA NECESSIDADE', col_data)] if col is None]
    if faltando:
        raise KeyError(f"Colunas obrigatórias não encontradas no Plano: {faltando}. Colunas disponíveis: {df_plano.columns.tolist()}")

    df_plano['cod_mat'] = df_plano[col_cod_sap].apply(normaliza_cod)
    df_plano['empresa']  = df_plano[col_empresa].astype(str).str.strip()
    df_plano['chave']    = df_plano['cod_mat'] + '|' + df_plano['empresa']
    df_plano['DATA NECESSIDADE'] = pd.to_datetime(df_plano[col_data], errors='coerce')
    df_plano['Mes_consumo']      = df_plano['DATA NECESSIDADE'].dt.to_period('M').dt.to_timestamp()
    df_plano['QTD ITEM']         = pd.to_numeric(df_plano[col_qtd], errors='coerce').fillna(0)
    if col_classe:
        df_plano['CLASSE MATERIAL'] = df_plano[col_classe]
    if col_subcls:
        df_plano['SUBCLASSE MATERIAL'] = df_plano[col_subcls]
    if col_blq:
        df_plano['BLQ?'] = df_plano[col_blq]

    if 'CLASSE MATERIAL' in df_plano.columns:
        excl = df_plano[df_plano['CLASSE MATERIAL'].astype(str).str.strip().isin(CLASSES_EXCLUIR)]
        print(f"  🚫 Plano — excluídos por CLASSE MATERIAL: {len(excl)} linhas")
    if 'SUBCLASSE MATERIAL' in df_plano.columns:
        excl = df_plano[df_plano['SUBCLASSE MATERIAL'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)]
        print(f"  🚫 Plano — excluídos por SUBCLASSE MATERIAL: {len(excl)} linhas")
    if 'BLQ?' in df_plano.columns:
        excl = df_plano[df_plano['BLQ?'].astype(str).str.upper().str.strip().isin(['SIM', 'BLOQUEADO'])]
        print(f"  🚫 Plano — excluídos por BLQ?: {len(excl)} linhas")

    # Metadata (classe/subclasse/descrição) capturada ANTES da exclusão,
    # para que materiais excluídos do consumo ainda apareçam identificáveis no relatório.
    meta_classe = df_plano.groupby('chave', as_index=False).agg(
        classe=('CLASSE MATERIAL','first'), subclasse=('SUBCLASSE MATERIAL','first')
    )

    if 'CLASSE MATERIAL' in df_plano.columns:
        df_plano = df_plano[~df_plano['CLASSE MATERIAL'].astype(str).str.strip().isin(CLASSES_EXCLUIR)].copy()
    if 'SUBCLASSE MATERIAL' in df_plano.columns:
        df_plano = df_plano[~df_plano['SUBCLASSE MATERIAL'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)].copy()
    if 'BLQ?' in df_plano.columns:
        df_plano = df_plano[~df_plano['BLQ?'].astype(str).str.upper().str.strip().isin(['SIM', 'BLOQUEADO'])].copy()

    df_plano = df_plano[
        (df_plano['QTD ITEM'] > 0) &
        (df_plano['Mes_consumo'] >= CORTE) &
        (df_plano['DATA NECESSIDADE'].dt.year == 2026)
    ].copy()

    consumo = df_plano.groupby(['chave', 'Mes_consumo'], as_index=False).agg(
        Qtd_consumo=('QTD ITEM', 'sum')
    ).rename(columns={'Mes_consumo': 'Mes'})

    print(f"  ✅ Plano irrestrito: {len(consumo)} consumos mensais")

    return df_est, entradas, consumo, meta_classe, df_ped


# ─────────────────────────────────────────────
# ETAPA 1 — PROJEÇÃO DE QUANTIDADE (sem PU)
# ─────────────────────────────────────────────
def projetar_quantidade(df_est, entradas, consumo, meta_classe):
    print("\n⚙️  Calculando projeção de QUANTIDADE (sem valorar ainda)...")

    saldo_qtd = df_est.set_index('chave')['Qtd_estoque'].to_dict()
    meta_cod  = df_est.set_index('chave')['cod'].to_dict()
    meta_emp  = df_est.set_index('chave')['empresa'].to_dict()
    meta_desc = df_est.set_index('chave')['descricao'].to_dict()
    meta_cls  = meta_classe.set_index('chave')['classe'].to_dict()
    meta_sub  = meta_classe.set_index('chave')['subclasse'].to_dict()

    ent_idx = entradas.set_index(['chave', 'Mes_entrega']) if not entradas.empty else pd.DataFrame()
    con_idx = consumo.set_index(['chave', 'Mes'])          if not consumo.empty  else pd.DataFrame()

    todas_chaves = (
        set(saldo_qtd.keys()) |
        set(entradas['chave'].unique()) |
        set(consumo['chave'].unique())
    )

    for chave in todas_chaves:
        if chave not in meta_cod:
            partes = chave.split('|')
            meta_cod[chave]  = partes[0] if partes else chave
            meta_emp[chave]  = partes[1] if len(partes) > 1 else ''
            meta_desc[chave] = ''
            meta_cls[chave]  = ''
            meta_sub[chave]  = ''

    linhas = []  # um dicionário por (material, mês) — só quantidade, sem R$

    for chave in sorted(todas_chaves):
        qtd = max(float(saldo_qtd.get(chave, 0.0)), 0.0)

        for mes in MESES:
            try:
                e = ent_idx.loc[(chave, mes)]
                qtd_ent = float(e['Qtd_entrada']) if isinstance(e, pd.Series) else float(e['Qtd_entrada'].sum())
            except KeyError:
                qtd_ent = 0.0

            try:
                c = con_idx.loc[(chave, mes)]
                qtd_con_solicitado = float(c['Qtd_consumo']) if isinstance(c, pd.Series) else float(c['Qtd_consumo'].sum())
            except KeyError:
                qtd_con_solicitado = 0.0

            qtd_disp  = qtd + qtd_ent
            qtd_con   = min(qtd_con_solicitado, qtd_disp)
            qtd_final = max(qtd_disp - qtd_con, 0.0)

            linhas.append({
                'chave':       chave,
                'cod':         meta_cod.get(chave, ''),
                'empresa':     meta_emp.get(chave, ''),
                'descricao':   meta_desc.get(chave, ''),
                'classe':      meta_cls.get(chave, ''),
                'subclasse':   meta_sub.get(chave, ''),
                'mes':         mes_ptbr(mes),
                'qtd_ini':     round(qtd, 2),
                'qtd_entrada': round(qtd_ent, 2),
                'qtd_consumo': round(qtd_con, 2),
                'qtd_final':   round(qtd_final, 2),
            })

            qtd = qtd_final

    return pd.DataFrame(linhas)


# ─────────────────────────────────────────────
# ETAPA 2 — MONTAGEM DO PU EFETIVO (isolada, defensiva)
# ─────────────────────────────────────────────
def montar_pu_efetivo(df_est, df_ped):
    """Monta o dicionário chave -> PU, combinando (em ordem de prioridade):
       1) custo médio real do estoque (Valor/Qtd)
       2) tabela oficial de PU (aba PU do arquivo de Plano)
       3) preço médio do próprio pedido (fallback final)
       Qualquer problema de leitura da tabela PU só reduz a cobertura —
       nunca quebra a execução."""

    pu_dict = {}

    # 1) Tenta ler a tabela oficial de PU
    try:
        df_pu, header_usado, col_cod = ler_aba_com_cabecalho_flexivel(
            ARQ_PLANO, ABA_PU,
            candidatos_coluna_chave=['COD', 'COD SAP', 'CÓD MATERIAL', 'Código', 'Material']
        )
        print(f"  🔎 PU — cabeçalho encontrado na linha {header_usado} (coluna de código: '{col_cod}')")

        col_emp = encontra_coluna(df_pu, ['Empresa', 'EMPRESA'])
        col_pu  = encontra_coluna(df_pu, ['PU', 'Preço Unitário', 'Preco Unitario', 'Unit'])

        if col_emp and col_pu:
            df_pu['COD_str']  = df_pu[col_cod].apply(normaliza_cod)
            df_pu['chave_pu'] = df_pu['COD_str'] + '|' + df_pu[col_emp].astype(str).str.strip()
            df_pu['Unit_num'] = pd.to_numeric(df_pu[col_pu], errors='coerce').fillna(0)
            pu_dict = df_pu.set_index('chave_pu')['Unit_num'].to_dict()
            print(f"  ✅ Tabela PU: {len(pu_dict)} preços carregados (colunas: {col_cod} / {col_emp} / {col_pu})")
        else:
            print(f"  ⚠️  Tabela PU: não encontrou empresa/PU (empresa={col_emp}, pu={col_pu}). Colunas disponíveis: {df_pu.columns.tolist()}")
            print(f"  ⚠️  Seguindo sem a tabela oficial de PU — será usado custo médio do estoque e preço do pedido.")
    except Exception as ex:
        print(f"  ⚠️  Não foi possível ler a aba '{ABA_PU}' em '{ARQ_PLANO.name}': {ex}")
        print(f"  ⚠️  Seguindo sem a tabela oficial de PU — será usado custo médio do estoque e preço do pedido.")

    # 2) Custo médio real do estoque tem prioridade — sobrescreve o PU oficial
    substituidos, completados = 0, 0
    for _, row in df_est.iterrows():
        if row['Qtd_estoque'] > 0:
            pu_real = row['Valor_estoque'] / row['Qtd_estoque']
            if row['chave'] in pu_dict:
                substituidos += 1
            else:
                completados += 1
            pu_dict[row['chave']] = pu_real
    print(f"  ✅ PU efetivo: {substituidos} substituídos pelo custo médio real + {completados} completados (sem PU oficial)")

    # 3) Fallback final: preço médio do próprio pedido, para chaves ainda sem PU
    if 'Qtd_pedido' in df_ped.columns and 'Val_pedido' in df_ped.columns:
        chaves_pedido = set(df_ped.loc[df_ped['Qtd_pedido'] > 0, 'chave'].unique())
        sem_pu = chaves_pedido - set(pu_dict.keys())
        completados_via_pedido = 0
        for chave in sem_pu:
            sub = df_ped[(df_ped['chave'] == chave) & (df_ped['Qtd_pedido'] > 0)]
            if not sub.empty and sub['Qtd_pedido'].sum() > 0:
                pu_dict[chave] = sub['Val_pedido'].sum() / sub['Qtd_pedido'].sum()
                completados_via_pedido += 1
        print(f"  ✅ PU completado via preço do próprio pedido: {completados_via_pedido} chaves")

    return pu_dict


# ─────────────────────────────────────────────
# ETAPA 3 — VALORAÇÃO (aplica PU sobre as quantidades já prontas)
# ─────────────────────────────────────────────
def valorar(df_qtd, pu_dict):
    print("\n💰 Valorando as quantidades projetadas...")

    df = df_qtd.copy()
    df['pu']         = df['chave'].map(pu_dict).fillna(0.0)
    df['rs_ini']      = (df['qtd_ini']     * df['pu']).round(2)
    df['rs_entrada']  = (df['qtd_entrada'] * df['pu']).round(2)
    df['rs_consumo']  = (df['qtd_consumo'] * df['pu']).round(2)
    df['rs_final']    = (df['qtd_final']   * df['pu']).round(2)

    sem_pu = df.loc[df['pu'] == 0, 'chave'].nunique()
    if sem_pu:
        print(f"  ⚠️  {sem_pu} materiais ficaram com PU = 0 (valor R$ zerado, mas quantidade projetada normalmente)")

    df = df.drop(columns='chave')
    return df


# ─────────────────────────────────────────────
# RESUMO MENSAL
# ─────────────────────────────────────────────
def resumo_mensal(df_proj):
    r = df_proj.groupby('mes', sort=False).agg(
        **{'Entradas (R$)':    ('rs_entrada', 'sum'),
           'Consumo (R$)':     ('rs_consumo', 'sum'),
           'Saldo final (R$)': ('rs_final',   'sum')}
    ).reset_index()

    ordem = [mes_ptbr(m) for m in MESES]
    r['_ord'] = r['mes'].map({m: i for i, m in enumerate(ordem)})
    r = r.sort_values('_ord').drop(columns='_ord')
    r.rename(columns={'mes': 'Mês'}, inplace=True)

    ini_por_mes = df_proj.groupby('mes')['rs_ini'].sum().to_dict()
    r['Estoque inicial (R$)'] = r['Mês'].map(ini_por_mes)
    r = r[['Mês', 'Estoque inicial (R$)', 'Entradas (R$)', 'Consumo (R$)', 'Saldo final (R$)']]

    return r


# ─────────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────────
def exportar(df_proj, df_resumo):
    print(f"\n💾 Salvando Excel em: {ARQ_SAIDA}")
    with pd.ExcelWriter(ARQ_SAIDA, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='RESUMO MENSAL',       index=False)
        df_proj.to_excel(writer,   sheet_name='DETALHE POR MATERIAL', index=False)
    print("✅ Excel salvo!")

    resumo_json = df_resumo.rename(columns={
        'Mês':'mes','Estoque inicial (R$)':'ini',
        'Entradas (R$)':'ent','Consumo (R$)':'con','Saldo final (R$)':'sal'
    }).to_dict(orient='records')
    with open(PASTA_BASE / 'projecao_2026.json', 'w', encoding='utf-8') as f:
        json.dump(resumo_json, f, ensure_ascii=False)
    print("✅ projecao_2026.json salvo!")

    detalhe_json = df_proj.to_dict(orient='records')
    with open(PASTA_BASE / 'detalhe_2026.json', 'w', encoding='utf-8') as f:
        json.dump(detalhe_json, f, ensure_ascii=False, default=str)
    print("✅ detalhe_2026.json salvo!")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def executar():
    df_est, entradas, consumo, meta_classe, df_ped = ler_bases()

    df_qtd  = projetar_quantidade(df_est, entradas, consumo, meta_classe)
    pu_dict = montar_pu_efetivo(df_est, df_ped)
    df_proj = valorar(df_qtd, pu_dict)

    df_resumo = resumo_mensal(df_proj)

    print("\n📊 RESUMO DA PROJEÇÃO 2026:")
    for _, row in df_resumo.iterrows():
        print(f"  {row['Mês']:>7}  ini={row['Estoque inicial (R$)']:>15,.0f}  ent={row['Entradas (R$)']:>15,.0f}  con={row['Consumo (R$)']:>15,.0f}  sal={row['Saldo final (R$)']:>15,.0f}")

    exportar(df_proj, df_resumo)


if __name__ == '__main__':
    executar()