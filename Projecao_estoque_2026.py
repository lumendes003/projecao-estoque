"""
PROJEÇÃO DE ESTOQUE FINANCEIRO - EQTL 2026
===========================================
Lógica: Preço Médio Ponderado mês a mês

  Estoque inicial = saldo real atualizado (base do dia)
  Entradas        = pedidos com entrega >= MES_INICIO
  Consumo         = plano irrestrito agrupado por COD SAP + EMPRESA + mês da DATA NECESSIDADE
  Saldo final     = Saldo R$ + Entradas R$ − Consumo R$

  ⚙️  Para atualizar todo mês: altere apenas MES_INICIO abaixo.
"""

from pathlib import Path
import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# CONFIGURAÇÃO — altere MES_INICIO todo mês
# ─────────────────────────────────────────────
MES_INICIO  = '2026-05-01'   # ← atualizar mensalmente

PASTA_BASE  = Path(r'C:/Users/u10952/COBERTURA_BI/PROJEÇÃO_ESTOQUE')
ARQ_ENTRADA = PASTA_BASE / 'entrada.xlsx'
ARQ_PLANO   = Path(r'\\10.7.1.90\Expansao_At_Automacao\9. EXEC_CONTROLE OBRAS AT\BASES\PLANO IRRESTRITO VERSÃO ATUAL BI.xlsm')
ARQ_SAIDA   = Path(r'\\10.7.1.90\Expansao_At_Automacao\9. EXEC_CONTROLE OBRAS AT\BASES\PROJECAO_ESTOQUE_FINANCEIRO_2026.xlsx')

ABA_ESTOQUE = 'Estoque'
ABA_PEDIDOS = 'PEDIDOS'
ABA_PLANO   = 'LISTA + RESERVAS'

TODOS_MESES = pd.date_range('2026-01-01', '2026-12-01', freq='MS')
CORTE       = pd.Timestamp(MES_INICIO)
MESES       = [m for m in TODOS_MESES if m >= CORTE]

def mes_ptbr(m):
    return m.strftime('%b/%y').upper()\
        .replace('MAY','MAI').replace('AUG','AGO')\
        .replace('SEP','SET').replace('OCT','OUT')\
        .replace('DEC','DEZ')


# ─────────────────────────────────────────────
# LEITURA
# ─────────────────────────────────────────────
def ler_bases():
    print(f"📂 Lendo bases... (projeção: {mes_ptbr(CORTE)} → DEZ/26)")

    # ── ESTOQUE ──────────────────────────────
    # Colunas esperadas no Excel:
    #   CÓD MATERIAL, EMPRESA, DEPÓSITO, LOTE, DESCRIÇÃO MATERIAL,
    #   QUANTIDADE, UM BÁSICA, VALOR, ELEMENTO PEP, INSTALAÇÃO
    df_est = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_ESTOQUE, header=0)
    df_est['chave'] = (
        df_est['CÓD MATERIAL'].astype(str).str.strip() + '|' +
        df_est['EMPRESA'].astype(str).str.strip()
    )
    # 'Material' e 'CÓD MATERIAL' são a mesma coluna — usando CÓD MATERIAL
    df_est['cod']       = df_est['CÓD MATERIAL'].astype(str).str.strip()
    df_est['empresa']   = df_est['EMPRESA'].astype(str).str.strip()
    # Descrição: coluna renomeada de 'Texto breve de material' → 'DESCRIÇÃO MATERIAL'
    df_est['descricao'] = df_est['DESCRIÇÃO MATERIAL'].astype(str).str.strip() \
                          if 'DESCRIÇÃO MATERIAL' in df_est.columns else ''

    # Quantidade: coluna renomeada de 'Quantidade de Material' → 'QUANTIDADE'
    df_est['Qtd_estoque']   = pd.to_numeric(df_est['QUANTIDADE'], errors='coerce').fillna(0)
    # Valor: mesmo nome, só caixa alta → 'VALOR'
    df_est['Valor_estoque'] = pd.to_numeric(df_est['VALOR'], errors='coerce').fillna(0)

    df_meta = df_est.groupby('chave', as_index=False).agg(
        cod       = ('cod',       'first'),
        empresa   = ('empresa',   'first'),
        descricao = ('descricao', 'first')
    )
    df_est = df_est.groupby('chave', as_index=False).agg(
        Qtd_estoque   = ('Qtd_estoque',   'sum'),
        Valor_estoque = ('Valor_estoque', 'sum')
    )
    df_est = df_est.merge(df_meta, on='chave', how='left')
    df_est['Preco_medio_ini'] = np.where(
        df_est['Qtd_estoque'] > 0,
        df_est['Valor_estoque'] / df_est['Qtd_estoque'],
        0.0
    )
    print(f"  ✅ Estoque: {len(df_est)} materiais")

    # ── PEDIDOS ───────────────────────────────
    # Colunas esperadas no Excel:
    #   EMPRESA, ReqC, Pedido, Item, Fornecedor/centro fornecedor, Material,
    #   Texto breve, Soma de Qtd.do pedido, Qtd.a fornecer, Valor,
    #   Data do documento, Data de remessa, Dat.rem.estatística, PEP, INSTALAÇÃO
    df_ped = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_PEDIDOS, header=0)
    df_ped['chave'] = (
        df_ped['Material'].astype(str).str.strip() + '|' +
        df_ped['EMPRESA'].astype(str).str.strip()
    )
    # Data de entrega: coluna 'Dat.rem.estatística' — mesmo nome ✅
    df_ped['Data_entrega'] = pd.to_datetime(df_ped['Dat.rem.estatística'], dayfirst=True, errors='coerce')
    # Quantidade: coluna renomeada de 'a ser fornecida (quantidade)' → 'Qtd.a fornecer'
    df_ped['Qtd_pedido']   = pd.to_numeric(df_ped['Qtd.a fornecer'], errors='coerce').fillna(0)
    # Valor: coluna renomeada de 'a ser fornecido (valor)' → 'Valor'
    df_ped['Valor_total']  = pd.to_numeric(df_ped['Valor'], errors='coerce').fillna(0)
    df_ped['Mes_entrega']  = df_ped['Data_entrega'].dt.to_period('M').dt.to_timestamp()

    df_ped = df_ped[
        (df_ped['Qtd_pedido'] > 0) &
        (df_ped['Mes_entrega'] >= CORTE)
    ].copy()

    entradas = df_ped.groupby(['chave', 'Mes_entrega'], as_index=False).agg(
        Qtd_entrada   = ('Qtd_pedido',  'sum'),
        Valor_entrada = ('Valor_total', 'sum')
    )

    # Preço médio do pedido por chave (fallback nível 2)
    df_ped_pm = df_ped.groupby('chave', as_index=False).agg(
        Qtd_ped_total = ('Qtd_pedido', 'sum'),
        Val_ped_total = ('Valor_total', 'sum')
    )
    df_ped_pm['PM_ped'] = np.where(
        df_ped_pm['Qtd_ped_total'] > 0,
        df_ped_pm['Val_ped_total'] / df_ped_pm['Qtd_ped_total'],
        0.0
    )
    print(f"  ✅ Pedidos: {len(entradas)} entradas mensais por material (>= {mes_ptbr(CORTE)})")

    # ── PLANO IRRESTRITO ──────────────────────
    df_plano = pd.read_excel(ARQ_PLANO, sheet_name=ABA_PLANO, header=0, engine='openpyxl')
    df_plano['chave'] = (
        df_plano['COD SAP'].astype(str).str.strip() + '|' +
        df_plano['EMPRESA'].astype(str).str.strip()
    )
    df_plano['DATA NECESSIDADE'] = pd.to_datetime(df_plano['DATA NECESSIDADE'], errors='coerce')
    df_plano['Mes_consumo']      = df_plano['DATA NECESSIDADE'].dt.to_period('M').dt.to_timestamp()
    df_plano['QTD ITEM']         = pd.to_numeric(df_plano['QTD ITEM'], errors='coerce').fillna(0)
    df_plano['PU']               = pd.to_numeric(df_plano['PU'], errors='coerce').fillna(0)

    # Exclui itens bloqueados
    if 'BLQ?' in df_plano.columns:
        df_plano = df_plano[df_plano['BLQ?'].astype(str).str.upper().str.strip() != 'SIM'].copy()

    # Classe e subclasse
    meta_classe = df_plano.groupby('chave', as_index=False).agg(
        classe    = ('CLASSE MATERIAL',    'first'),
        subclasse = ('SUBCLASSE MATERIAL', 'first')
    )

    # PU fallback por chave
    pu_plano = df_plano.groupby('chave', as_index=False).agg(
        PU_medio = ('PU', 'mean')
    )

    # Filtra apenas meses >= corte e ano 2026
    df_plano = df_plano[
        (df_plano['QTD ITEM'] > 0) &
        (df_plano['Mes_consumo'] >= CORTE) &
        (df_plano['DATA NECESSIDADE'].dt.year == 2026)
    ].copy()

    df_consumo = df_plano.groupby(['chave', 'Mes_consumo'], as_index=False).agg(
        Qtd_consumo = ('QTD ITEM', 'sum')
    ).rename(columns={'Mes_consumo': 'Mes'})

    print(f"  ✅ Plano irrestrito: {len(df_consumo)} consumos mensais por material (>= {mes_ptbr(CORTE)})")

    return df_est, entradas, df_consumo, meta_classe, df_ped_pm, pu_plano


# ─────────────────────────────────────────────
# PROJEÇÃO MÊS A MÊS
# ─────────────────────────────────────────────
def projetar(df_est, df_entradas, df_consumo, meta_classe, df_ped_pm, pu_plano):
    print("\n⚙️  Calculando projeção com preço médio ponderado...")

    pm_ped_dict = df_ped_pm.set_index('chave')['PM_ped'].to_dict()
    pu_dict     = pu_plano.set_index('chave')['PU_medio'].to_dict()

    saldo_qtd = df_est.set_index('chave')['Qtd_estoque'].to_dict()
    saldo_rs  = df_est.set_index('chave')['Valor_estoque'].to_dict()
    meta_cod  = df_est.set_index('chave')['cod'].to_dict()
    meta_emp  = df_est.set_index('chave')['empresa'].to_dict()
    meta_desc = df_est.set_index('chave')['descricao'].to_dict()
    meta_cls  = meta_classe.set_index('chave')['classe'].to_dict()
    meta_sub  = meta_classe.set_index('chave')['subclasse'].to_dict()

    ent_idx = (
        df_entradas.set_index(['chave', 'Mes_entrega'])
        if not df_entradas.empty else pd.DataFrame()
    )
    con_idx = (
        df_consumo.set_index(['chave', 'Mes'])
        if not df_consumo.empty else pd.DataFrame()
    )

    todas_chaves = (
        set(saldo_qtd.keys()) |
        set(df_entradas['chave'].unique()) |
        set(df_consumo['chave'].unique())
    )

    for chave in todas_chaves:
        if chave not in meta_cod:
            partes = chave.split('|')
            meta_cod[chave]  = partes[0] if len(partes) > 0 else chave
            meta_emp[chave]  = partes[1] if len(partes) > 1 else ''
            meta_desc[chave] = ''
            meta_cls[chave]  = ''
            meta_sub[chave]  = ''

    linhas = []

    for chave in sorted(todas_chaves):
        qtd = float(saldo_qtd.get(chave, 0.0))
        rs  = float(saldo_rs.get(chave,  0.0))

        for mes in MESES:

            try:
                e = ent_idx.loc[(chave, mes)]
                qtd_ent = float(e['Qtd_entrada'])  if isinstance(e, pd.Series) else float(e['Qtd_entrada'].sum())
                rs_ent  = float(e['Valor_entrada']) if isinstance(e, pd.Series) else float(e['Valor_entrada'].sum())
            except KeyError:
                qtd_ent, rs_ent = 0.0, 0.0

            qtd_disp    = qtd + qtd_ent
            rs_disp     = rs  + rs_ent
            if qtd_disp > 0:
                preco_medio = rs_disp / qtd_disp
            elif pm_ped_dict.get(chave, 0.0) > 0:
                preco_medio = pm_ped_dict[chave]
            else:
                preco_medio = pu_dict.get(chave, 0.0)

            try:
                c = con_idx.loc[(chave, mes)]
                qtd_con = float(c['Qtd_consumo']) if isinstance(c, pd.Series) else float(c['Qtd_consumo'].sum())
            except KeyError:
                qtd_con = 0.0

            rs_con    = qtd_con * preco_medio
            qtd_final = qtd_disp - qtd_con
            rs_final  = rs_disp  - rs_con

            linhas.append({
                'cod':         meta_cod.get(chave, ''),
                'empresa':     meta_emp.get(chave, ''),
                'descricao':   meta_desc.get(chave, ''),
                'classe':      meta_cls.get(chave, ''),
                'subclasse':   meta_sub.get(chave, ''),
                'mes':         mes_ptbr(mes),
                'qtd_ini':     round(qtd, 2),
                'rs_ini':      round(rs, 2),
                'qtd_entrada': round(qtd_ent, 2),
                'rs_entrada':  round(rs_ent, 2),
                'preco_medio': round(preco_medio, 4),
                'qtd_consumo': round(qtd_con, 2),
                'rs_consumo':  round(rs_con, 2),
                'qtd_final':   round(qtd_final, 2),
                'rs_final':    round(rs_final, 2),
            })

            qtd, rs = qtd_final, rs_final

    return pd.DataFrame(linhas)


# ─────────────────────────────────────────────
# RESUMO MENSAL
# ─────────────────────────────────────────────
def resumo_mensal(df_proj):
    r = df_proj.groupby('mes', sort=False).agg(
        **{'Estoque inicial (R$)': ('rs_ini',     'sum'),
           'Entradas (R$)':        ('rs_entrada', 'sum'),
           'Consumo (R$)':         ('rs_consumo', 'sum'),
           'Saldo final (R$)':     ('rs_final',   'sum')}
    ).reset_index()

    ordem = [mes_ptbr(m) for m in MESES]
    r['_ord'] = r['mes'].map({m: i for i, m in enumerate(ordem)})
    r = r.sort_values('_ord').drop(columns='_ord')
    r.rename(columns={'mes': 'Mês'}, inplace=True)
    return r


# ─────────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────────
def exportar(df_proj, df_resumo):
    print(f"\n💾 Salvando Excel em: {ARQ_SAIDA}")
    with pd.ExcelWriter(ARQ_SAIDA, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='RESUMO MENSAL',        index=False)
        df_proj.to_excel(writer,   sheet_name='DETALHE POR MATERIAL',  index=False)
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
    df_est, df_entradas, df_consumo, meta_classe, df_ped_pm, pu_plano = ler_bases()
    df_proj   = projetar(df_est, df_entradas, df_consumo, meta_classe, df_ped_pm, pu_plano)
    df_resumo = resumo_mensal(df_proj)

    print("\n📊 RESUMO DA PROJEÇÃO 2026:")
    for _, row in df_resumo.iterrows():
        print(f"  {row['Mês']:>7}  ini={row['Estoque inicial (R$)']:>15,.0f}  ent={row['Entradas (R$)']:>15,.0f}  con={row['Consumo (R$)']:>15,.0f}  sal={row['Saldo final (R$)']:>15,.0f}")

    exportar(df_proj, df_resumo)


if __name__ == '__main__':
    executar()