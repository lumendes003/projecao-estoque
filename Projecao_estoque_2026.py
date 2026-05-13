"""
PROJEÇÃO DE ESTOQUE FINANCEIRO - EQTL 2026
===========================================
Lógica: Preço Médio Ponderado mês a mês

  Estoque inicial = saldo real atualizado (base do dia)
  Entradas        = pedidos com entrega >= MES_INICIO
  Consumo         = plano a partir de MES_INICIO
  Saldo final     = Saldo R$ + Entradas R$ − Consumo R$

  ⚙️  Para atualizar todo mês: altere apenas MES_INICIO abaixo.

  Obs: coluna 'Valor' dos pedidos = valor total do item (qtd × preço unitário).
       Preço unitário = Valor / Qtd.a fornecer.
"""

from pathlib import Path
import pandas as pd
import numpy as np
import json

# ─────────────────────────────────────────────
# CONFIGURAÇÃO — altere MES_INICIO todo mês
# ─────────────────────────────────────────────
MES_INICIO  = '2026-05-01'   # ← atualizar mensalmente

PASTA_BASE  = Path(r'C:/Users/u10952/COBERTURA_BI/PROJEÇÃO_ESTOQUE')
ARQ_ENTRADA = PASTA_BASE / 'ENTRADA.xlsx'
ARQ_SAIDA   = PASTA_BASE / 'PROJECAO_ESTOQUE_FINANCEIRO_2026.xlsx'

ABA_ESTOQUE = 'Estoque'
ABA_PEDIDOS = 'PEDIDOS'
ABA_PLANO   = 'PLANO 2026'

SIGLAS      = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ']
TODOS_MESES = pd.date_range('2026-01-01', '2026-12-01', freq='MS')

CORTE = pd.Timestamp(MES_INICIO)
MESES = [m for m in TODOS_MESES if m >= CORTE]

COLUNAS_QTD_DI = [f'Soma de {s} DI'    for s in SIGLAS]
COLUNAS_RS_DI  = [f'Soma de {s} DI R$' for s in SIGLAS]


# ─────────────────────────────────────────────
# LEITURA
# ─────────────────────────────────────────────
def ler_bases():
    print(f"📂 Lendo bases... (projeção: {CORTE.strftime('%b/%y').upper()} → DEZ/26)")

    # ── ESTOQUE ──────────────────────────────
    df_est = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_ESTOQUE, header=0)
    df_est['chave'] = (
        df_est['CÓD MATERIAL'].astype(str).str.strip() +
        df_est['EMPRESA'].astype(str).str.strip()
    )
    df_est['Qtd_estoque']   = pd.to_numeric(df_est['QUANTIDADE'], errors='coerce').fillna(0)
    df_est['Valor_estoque'] = pd.to_numeric(df_est['VALOR'],      errors='coerce').fillna(0)

    df_est = df_est.groupby('chave', as_index=False).agg(
        Qtd_estoque   = ('Qtd_estoque',  'sum'),
        Valor_estoque = ('Valor_estoque', 'sum')
    )
    df_est['Preco_medio_ini'] = np.where(
        df_est['Qtd_estoque'] > 0,
        df_est['Valor_estoque'] / df_est['Qtd_estoque'],
        0.0
    )
    print(f"  ✅ Estoque: {len(df_est)} materiais")

    # ── PEDIDOS ───────────────────────────────
    df_ped = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_PEDIDOS, header=0)
    df_ped['chave'] = (
        df_ped['Material'].astype(str).str.strip() +
        df_ped['EMPRESA'].astype(str).str.strip()
    )
    df_ped['Data_entrega'] = pd.to_datetime(df_ped['Data de remessa'], dayfirst=True, errors='coerce')
    df_ped['Qtd_pedido']   = pd.to_numeric(df_ped['Qtd.a fornecer'], errors='coerce').fillna(0)

    # Valor já é o total do item — preço unitário = Valor / Qtd
    df_ped['Valor_total'] = pd.to_numeric(df_ped['Valor'], errors='coerce').fillna(0)
    df_ped['Preco_unit']  = np.where(
        df_ped['Qtd_pedido'] > 0,
        df_ped['Valor_total'] / df_ped['Qtd_pedido'],
        0.0
    )

    df_ped['Mes_entrega'] = df_ped['Data_entrega'].dt.to_period('M').dt.to_timestamp()

    # Filtra apenas entregas a partir do corte e com quantidade pendente
    df_ped = df_ped[
        (df_ped['Qtd_pedido'] > 0) &
        (df_ped['Mes_entrega'] >= CORTE)
    ].copy()

    # Valor_entrada = valor total já está pronto (não multiplica de novo)
    entradas = df_ped.groupby(['chave', 'Mes_entrega'], as_index=False).agg(
        Qtd_entrada   = ('Qtd_pedido',   'sum'),
        Valor_entrada = ('Valor_total',   'sum')   # ← soma dos valores totais por mês
    )
    print(f"  ✅ Pedidos: {len(entradas)} entradas mensais por material (>= {CORTE.strftime('%b/%y').upper()})")

    # ── PLANO ─────────────────────────────────
    df_plano = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_PLANO, header=0)
    df_plano['chave'] = (
        df_plano['COD'].astype(str).str.strip() +
        df_plano['EMPRESA'].astype(str).str.strip()
    )

    for col in COLUNAS_QTD_DI + COLUNAS_RS_DI:
        if col in df_plano.columns:
            df_plano[col] = pd.to_numeric(df_plano[col], errors='coerce').fillna(0)
        else:
            df_plano[col] = 0.0

    consumo_rows = []
    for _, row in df_plano.iterrows():
        chave = row['chave']
        for i, mes in enumerate(TODOS_MESES):
            if mes < CORTE:
                continue
            qtd = row[COLUNAS_QTD_DI[i]]
            if qtd > 0:
                consumo_rows.append({
                    'chave':       chave,
                    'Mes':         mes,
                    'Qtd_consumo': qtd
                })

    df_consumo = pd.DataFrame(consumo_rows)
    if not df_consumo.empty:
        df_consumo = df_consumo.groupby(['chave', 'Mes'], as_index=False).agg(
            Qtd_consumo=('Qtd_consumo', 'sum')
        )
    print(f"  ✅ Plano: {len(df_consumo)} consumos mensais por material (>= {CORTE.strftime('%b/%y').upper()})")

    return df_est, entradas, df_consumo


# ─────────────────────────────────────────────
# PROJEÇÃO MÊS A MÊS
# ─────────────────────────────────────────────
def projetar(df_est, df_entradas, df_consumo):
    print("\n⚙️  Calculando projeção com preço médio ponderado...")

    saldo_qtd = df_est.set_index('chave')['Qtd_estoque'].to_dict()
    saldo_rs  = df_est.set_index('chave')['Valor_estoque'].to_dict()

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

    linhas = []

    for chave in sorted(todas_chaves):
        qtd = float(saldo_qtd.get(chave, 0.0))
        rs  = float(saldo_rs.get(chave,  0.0))

        for mes in MESES:

            # Entradas
            try:
                e = ent_idx.loc[(chave, mes)]
                qtd_ent = float(e['Qtd_entrada'])  if isinstance(e, pd.Series) else float(e['Qtd_entrada'].sum())
                rs_ent  = float(e['Valor_entrada']) if isinstance(e, pd.Series) else float(e['Valor_entrada'].sum())
            except KeyError:
                qtd_ent, rs_ent = 0.0, 0.0

            # Preço médio ponderado
            qtd_disp    = qtd + qtd_ent
            rs_disp     = rs  + rs_ent
            preco_medio = rs_disp / qtd_disp if qtd_disp > 0 else 0.0

            # Consumo valorado pelo preço médio
            try:
                c = con_idx.loc[(chave, mes)]
                qtd_con = float(c['Qtd_consumo']) if isinstance(c, pd.Series) else float(c['Qtd_consumo'].sum())
            except KeyError:
                qtd_con = 0.0

            rs_con = qtd_con * preco_medio

            # Saldo final
            qtd_final = max(qtd_disp - qtd_con, 0.0)
            rs_final  = max(rs_disp  - rs_con,  0.0)

            linhas.append({
                'Chave':       chave,
                'Mês':         mes.strftime('%b/%y').upper(),
                'Qtd_ini':     round(qtd, 4),
                'RS_ini':      round(rs, 2),
                'Qtd_entrada': round(qtd_ent, 4),
                'RS_entrada':  round(rs_ent, 2),
                'Preço_médio': round(preco_medio, 4),
                'Qtd_consumo': round(qtd_con, 4),
                'RS_consumo':  round(rs_con, 2),
                'Qtd_final':   round(qtd_final, 4),
                'RS_final':    round(rs_final, 2),
            })

            qtd, rs = qtd_final, rs_final

    return pd.DataFrame(linhas)


# ─────────────────────────────────────────────
# RESUMO MENSAL (visão BI)
# ─────────────────────────────────────────────
def resumo_mensal(df_proj):
    r = df_proj.groupby('Mês', sort=False).agg(
        **{'Estoque inicial (R$)': ('RS_ini',     'sum'),
           'Entradas (R$)':        ('RS_entrada', 'sum'),
           'Consumo (R$)':         ('RS_consumo', 'sum'),
           'Saldo final (R$)':     ('RS_final',   'sum')}
    ).reset_index()

    ordem = [m.strftime('%b/%y').upper() for m in MESES]
    r['_ord'] = r['Mês'].map({m: i for i, m in enumerate(ordem)})
    r = r.sort_values('_ord').drop(columns='_ord')

    # Formata para leitura no terminal
    for col in ['Estoque inicial (R$)', 'Entradas (R$)', 'Consumo (R$)', 'Saldo final (R$)']:
        r[col] = r[col].apply(lambda v: f'R$ {v:,.0f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))

    return r


# ─────────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────────
def exportar(df_proj):
    print(f"\n💾 Salvando em: {ARQ_SAIDA}")

    df_resumo_excel = df_proj.groupby('Mês', sort=False).agg(
        **{'Estoque inicial (R$)': ('RS_ini',     'sum'),
           'Entradas (R$)':        ('RS_entrada', 'sum'),
           'Consumo (R$)':         ('RS_consumo', 'sum'),
           'Saldo final (R$)':     ('RS_final',   'sum')}
    ).reset_index()
    ordem = [m.strftime('%b/%y').upper() for m in MESES]
    df_resumo_excel['_ord'] = df_resumo_excel['Mês'].map({m: i for i, m in enumerate(ordem)})
    df_resumo_excel = df_resumo_excel.sort_values('_ord').drop(columns='_ord')
    with pd.ExcelWriter(ARQ_SAIDA, engine='openpyxl') as writer:
        df_resumo_excel.to_excel(writer, sheet_name='RESUMO MENSAL',        index=False)
        df_proj.to_excel(writer,         sheet_name='DETALHE POR MATERIAL',  index=False)    

    print("✅ Arquivo salvo com sucesso!")
    resumo_json = df_resumo_excel.rename(columns={
        'Mês':                 'mes',
        'Estoque inicial (R$)':'ini',
        'Entradas (R$)':       'ent',
        'Consumo (R$)':        'con',
        'Saldo final (R$)':    'sal'
    }).to_dict(orient='records')

    with open(PASTA_BASE / 'projecao_2026.json', 'w', encoding='utf-8') as f:
        json.dump(resumo_json, f, ensure_ascii=False)

    print("✅ JSON salvo: projecao_2026.json")
    


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def executar():
    df_est, df_entradas, df_consumo = ler_bases()
    df_proj   = projetar(df_est, df_entradas, df_consumo)
    df_resumo = resumo_mensal(df_proj)

    print("\n📊 RESUMO DA PROJEÇÃO 2026:")
    print(df_resumo.to_string(index=False))

    exportar(df_proj)


if __name__ == '__main__':
    executar()