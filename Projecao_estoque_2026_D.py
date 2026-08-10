"""
PROJEÇÃO DE ESTOQUE FINANCEIRO - EQTL 2026
Base: DILIGENCIAMENTO (Previsão de entrega)
"""

from pathlib import Path
import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

MES_INICIO  = '2026-08-01'

PASTA_BASE  = Path(r'\\10.7.1.90\Expansao_At_Automacao\9. EXEC_CONTROLE OBRAS AT\BASES\PROJEÇÃO_ESTOQUE')
ARQ_PLANO   = Path(r'\\10.7.1.90\Expansao_At_Automacao\5.EXEC_GESTÃO DE PROJETOS_AT\5 - Aquisição de Materiais\3 - Solicitações de Compra\3 Compra Prévia\ATUAL BI\COBERTURA POR OBRAS\BASE ENTRADA COBERTURA OBRAS.xlsm')
ARQ_ENTRADA = PASTA_BASE / 'entrada.xlsx'

ABA_ESTOQUE = 'Estoque'
ABA_PEDIDOS = 'DILIGENCIAMENTO'
ABA_PLANO   = 'PLANO 2025'
ABA_PU      = 'PLANO 2025'

ARQ_SAIDA   = PASTA_BASE / 'PROJECAO_ESTOQUE_FINANCEIRO_2026_D.xlsx'

SUBCLASSES_EXCLUIR = ['TORRE METÁLICA', 'TRANSFORMADOR DE FORÇA', 'MÓDULO GIS', 'REGULADOR DE TENSÃO']
CLASSES_EXCLUIR    = ['PRÉ-MOLDADO', 'TRANSFORMADOR DE FORÇA']

TODOS_MESES = pd.date_range('2026-01-01', '2026-12-01', freq='MS')
CORTE       = pd.Timestamp(MES_INICIO)
MESES       = [m for m in TODOS_MESES if m >= CORTE]

def mes_ptbr(m):
    return m.strftime('%b/%y').upper()\
        .replace('MAY','MAI').replace('AUG','AGO')\
        .replace('SEP','SET').replace('OCT','OUT')\
        .replace('DEC','DEZ')

def ler_bases():
    print(f"📂 Lendo bases DILIGENCIAMENTO... ({mes_ptbr(CORTE)} → DEZ/26)")

    # ── ESTOQUE ──────────────────────────────
    df_est = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_ESTOQUE, header=0)
    df_est['chave']     = df_est['CÓD MATERIAL'].astype(str).str.strip() + '|' + df_est['EMPRESA'].astype(str).str.strip()
    df_est['cod']       = df_est['CÓD MATERIAL'].astype(str).str.strip()
    df_est['empresa']   = df_est['EMPRESA'].astype(str).str.strip()
    df_est['descricao'] = df_est['DESCRIÇÃO MATERIAL'].astype(str).str.strip() if 'DESCRIÇÃO MATERIAL' in df_est.columns else ''
    df_est['Qtd_estoque']   = pd.to_numeric(df_est['QUANTIDADE'], errors='coerce').fillna(0)
    df_est['Valor_estoque'] = pd.to_numeric(df_est['VALOR'],      errors='coerce').fillna(0)

    df_meta = df_est.groupby('chave', as_index=False).agg(cod=('cod','first'), empresa=('empresa','first'), descricao=('descricao','first'))
    df_est  = df_est.groupby('chave', as_index=False).agg(Qtd_estoque=('Qtd_estoque','sum'), Valor_estoque=('Valor_estoque','sum'))
    df_est  = df_est.merge(df_meta, on='chave', how='left')
    print(f"  ✅ Estoque: {len(df_est)} materiais — R$ {df_est['Valor_estoque'].sum():,.0f}")

    # ── TABELA PU ────────────────────────────
    df_pu = pd.read_excel(ARQ_PLANO, sheet_name=ABA_PU, header=2)
    df_pu['COD_str']  = df_pu['COD SAP'].astype(str).str.strip()
    df_pu['chave_pu'] = df_pu['COD_str'] + '|' + df_pu['EMPRESA'].astype(str).str.strip()
    df_pu['Unit_num'] = pd.to_numeric(df_pu['PU'], errors='coerce').fillna(0)
    pu_dict = df_pu.set_index('chave_pu')['Unit_num'].to_dict()
    print(f"  ✅ Tabela PU: {len(pu_dict)} preços carregados")

    # ── DILIGENCIAMENTO ───────────────────────
    df_ped = pd.read_excel(ARQ_ENTRADA, sheet_name=ABA_PEDIDOS, header=0)
    df_ped['chave']      = df_ped['CÓDIGO'].astype(str).str.strip() + '|' + df_ped['EMPRESA'].astype(str).str.strip()
    df_ped['Qtd_pedido'] = pd.to_numeric(df_ped['QUANTIDADE'], errors='coerce').fillna(0)
    df_ped['Val_pedido'] = pd.to_numeric(df_ped['VALOR'],      errors='coerce').fillna(0)

    # Usa Previsão de entrega como data
    df_ped['Mes_entrega'] = pd.to_datetime(df_ped['Previsão de entrega'], errors='coerce').dt.to_period('M').dt.to_timestamp()

    # Filtros por classe e subclasse
    df_ped = df_ped[~df_ped['CLASSE MATERIAL'].astype(str).str.strip().isin(CLASSES_EXCLUIR)].copy()
    df_ped = df_ped[~df_ped['SUBCLASSE MATERIAL'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)].copy()

    # Só pedidos com qtd > 0 e entrega em 2026 >= corte
    df_ped = df_ped[
        (df_ped['Qtd_pedido'] > 0) &
        (df_ped['Mes_entrega'] >= CORTE) &
        (df_ped['Mes_entrega'].dt.year == 2026)
    ].copy()

    entradas = df_ped.groupby(['chave', 'Mes_entrega'], as_index=False).agg(
        Qtd_entrada=('Qtd_pedido', 'sum'),
        Val_entrada=('Val_pedido', 'sum')
    )
    print(f"  ✅ Diligenciamento: {len(entradas)} entradas mensais — R$ {df_ped['Val_pedido'].sum():,.0f}")

    # ── PLANO IRRESTRITO ──────────────────────
    df_plano = pd.read_excel(ARQ_PLANO, sheet_name=ABA_PLANO, header=2, engine='openpyxl')
    df_plano['chave'] = df_plano['COD SAP'].astype(str).str.strip() + '|' + df_plano['EMPRESA'].astype(str).str.strip()
    df_plano['DATA NECESSIDADE'] = pd.to_datetime(df_plano['DATA NECESSIDADE'], errors='coerce')
    df_plano['Mes_consumo']      = df_plano['DATA NECESSIDADE'].dt.to_period('M').dt.to_timestamp()
    df_plano['QTD ITEM']         = pd.to_numeric(df_plano['QTD ITEM'], errors='coerce').fillna(0)
    
    if 'CLASSE MATERIAL' in df_plano.columns:
            df_plano = df_plano[~df_plano['CLASSE MATERIAL'].astype(str).str.strip().isin(CLASSES_EXCLUIR)].copy()
    if 'SUBCLASSE MATERIAL' in df_plano.columns:
            df_plano = df_plano[~df_plano['SUBCLASSE MATERIAL'].astype(str).str.strip().isin(SUBCLASSES_EXCLUIR)].copy()
    if 'BLQ?' in df_plano.columns:
            df_plano = df_plano[~df_plano['BLQ?'].astype(str).str.upper().str.strip().isin(['SIM', 'BLOQUEADO'])].copy()
    
    meta_classe = df_plano.groupby('chave', as_index=False).agg(
        classe=('CLASSE MATERIAL','first'), subclasse=('SUBCLASSE MATERIAL','first')
    )

    df_plano = df_plano[
        (df_plano['QTD ITEM'] > 0) &
        (df_plano['Mes_consumo'] >= CORTE) &
        (df_plano['DATA NECESSIDADE'].dt.year == 2026)
    ].copy()

    consumo = df_plano.groupby(['chave', 'Mes_consumo'], as_index=False).agg(
        Qtd_consumo=('QTD ITEM', 'sum')
    ).rename(columns={'Mes_consumo': 'Mes'})

    print(f"  ✅ Plano irrestrito: {len(consumo)} consumos mensais")

    return df_est, entradas, consumo, meta_classe, pu_dict


def projetar(df_est, entradas, consumo, meta_classe, pu_dict):
    print("\n⚙️  Calculando projeção...")

    saldo_qtd = df_est.set_index('chave')['Qtd_estoque'].to_dict()
    saldo_val = df_est.set_index('chave')['Valor_estoque'].to_dict()
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

    linhas = []

    for chave in sorted(todas_chaves):
        qtd = float(saldo_qtd.get(chave, 0.0))
        pu  = pu_dict.get(chave, 0.0)

        for i, mes in enumerate(MESES):
            try:
                e = ent_idx.loc[(chave, mes)]
                qtd_ent = float(e['Qtd_entrada']) if isinstance(e, pd.Series) else float(e['Qtd_entrada'].sum())
                rs_ent  = float(e['Val_entrada'])  if isinstance(e, pd.Series) else float(e['Val_entrada'].sum())
            except KeyError:
                qtd_ent, rs_ent = 0.0, 0.0

            try:
                c = con_idx.loc[(chave, mes)]
                qtd_con = float(c['Qtd_consumo']) if isinstance(c, pd.Series) else float(c['Qtd_consumo'].sum())
            except KeyError:
                qtd_con = 0.0

            qtd_disp  = qtd + qtd_ent
            qtd_final = qtd_disp - qtd_con
            if i == 0:
                rs_ini = float(saldo_val.get(chave, 0.0))   # ← usa valor real no 1º mês
            else:
                rs_ini = max(qtd, 0.0) * pu

            
            rs_con   = max(qtd_con, 0.0)   * pu
            rs_final = max(qtd_final, 0.0) * pu

            linhas.append({
                'cod':         meta_cod.get(chave, ''),
                'empresa':     meta_emp.get(chave, ''),
                'descricao':   meta_desc.get(chave, ''),
                'classe':      meta_cls.get(chave, ''),
                'subclasse':   meta_sub.get(chave, ''),
                'mes':         mes_ptbr(mes),
                'pu':          round(pu, 4),
                'qtd_ini':     round(qtd, 2),
                'rs_ini':      round(rs_ini, 2),
                'qtd_entrada': round(qtd_ent, 2),
                'rs_entrada':  round(rs_ent, 2),
                'qtd_consumo': round(qtd_con, 2),
                'rs_consumo':  round(rs_con, 2),
                'qtd_final':   round(qtd_final, 2),
                'rs_final':    round(rs_final, 2),
            })

            qtd = qtd_final

    return pd.DataFrame(linhas)


def resumo_mensal(df_proj):
    r = df_proj.groupby('mes', sort=False).agg(
        **{'Entradas (R$)': ('rs_entrada','sum'), 'Consumo (R$)': ('rs_consumo','sum'), 'Saldo final (R$)': ('rs_final','sum')}
    ).reset_index()
    ordem = [mes_ptbr(m) for m in MESES]
    r['_ord'] = r['mes'].map({m: i for i, m in enumerate(ordem)})
    r = r.sort_values('_ord').drop(columns='_ord')
    r.rename(columns={'mes': 'Mês'}, inplace=True)
    ini_por_mes = df_proj.groupby('mes')['rs_ini'].sum().to_dict()
    r['Estoque inicial (R$)'] = r['Mês'].map(ini_por_mes)
    return r[['Mês', 'Estoque inicial (R$)', 'Entradas (R$)', 'Consumo (R$)', 'Saldo final (R$)']]


def exportar(df_proj, df_resumo):
    print(f"\n💾 Salvando Excel em: {ARQ_SAIDA}")
    with pd.ExcelWriter(ARQ_SAIDA, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='RESUMO MENSAL',       index=False)
        df_proj.to_excel(writer,   sheet_name='DETALHE POR MATERIAL', index=False)
    print("✅ Excel salvo!")

    resumo_json = df_resumo.rename(columns={
        'Mês':'mes','Estoque inicial (R$)':'ini','Entradas (R$)':'ent','Consumo (R$)':'con','Saldo final (R$)':'sal'
    }).to_dict(orient='records')
    with open(PASTA_BASE / 'projecao_2026_D.json', 'w', encoding='utf-8') as f:
        json.dump(resumo_json, f, ensure_ascii=False)
    print("✅ projecao_2026_D.json salvo!")

    detalhe_json = df_proj.to_dict(orient='records')
    with open(PASTA_BASE / 'detalhe_2026_D.json', 'w', encoding='utf-8') as f:
        json.dump(detalhe_json, f, ensure_ascii=False, default=str)
    print("✅ detalhe_2026_D.json salvo!")


def executar():
    df_est, entradas, consumo, meta_classe, pu_dict = ler_bases()
    df_proj   = projetar(df_est, entradas, consumo, meta_classe, pu_dict)
    df_resumo = resumo_mensal(df_proj)

    print("\n📊 RESUMO DA PROJEÇÃO 2026 (DILIGENCIAMENTO):")
    for _, row in df_resumo.iterrows():
        print(f"  {row['Mês']:>7}  ini={row['Estoque inicial (R$)']:>15,.0f}  ent={row['Entradas (R$)']:>15,.0f}  con={row['Consumo (R$)']:>15,.0f}  sal={row['Saldo final (R$)']:>15,.0f}")

    exportar(df_proj, df_resumo)

if __name__ == '__main__':
    executar()