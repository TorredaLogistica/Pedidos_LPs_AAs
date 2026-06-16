
import os
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

try:
    import matplotlib.pyplot as plt
    MATPLOT = True
except Exception:
    MATPLOT = False

ARQUIVO_PADRAO = 'Base OTIF SOE AAs e LPs.xlsx'
ABA = 'BASE OTIF'
TOP_LOJAS = 15
DIAS_MAP = {0: 'SEG', 1: 'TER', 2: 'QUA', 3: 'QUI', 4: 'SEX'}
ORDEM_DIAS = ['SEG', 'TER', 'QUA', 'QUI', 'SEX']
MESES_PT = {
    1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL', 5: 'MAIO', 6: 'JUNHO',
    7: 'JULHO', 8: 'AGOSTO', 9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'
}
DIA_PT = {
    0: 'segunda-feira', 1: 'terça-feira', 2: 'quarta-feira', 3: 'quinta-feira',
    4: 'sexta-feira', 5: 'sábado', 6: 'domingo'
}
MES_PT_EXT = {
    1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho',
    7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
}
REQ = [
    'PEDIDO', 'CD Origem', 'LOJA (SAP)', 'Data Criação', 'PROTOCOLO', 'DT-EXP',
    'Cód Cliente', 'Custo Faturamento', 'Tipo Centro', 'REG_DEST', 'UF_DEST', 'Semana da OV'
]

st.set_page_config(page_title='Indicadores Pedidos LPs', layout='wide')


def fmt_pct(v):
    try:
        return f'{float(v):.2f}%'.replace('.', ',')
    except Exception:
        return v


def fmt_num(v):
    try:
        return f'{float(v):.2f}'.replace('.', ',')
    except Exception:
        return v


def fmt_moeda(v):
    try:
        return f'R$ {float(v):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return v


def score_limite(v, limite, desconto=33.3333333333333):
    return 100.0 if float(v) <= float(limite) else 0.0


def score_excedente(exc, desconto=33.3333333333333):
    return 100.0 if float(exc) <= 0 else 0.0


def data_extenso(dt):
    if pd.isna(dt):
        return ''
    return f"{DIA_PT[dt.weekday()]}, {dt.day} de {MES_PT_EXT[dt.month]} de {dt.year}"

def render_metric_centered(title, value, title_size=16, value_size=34, margin_bottom=0):
    st.markdown(
        f"""
        <div style='text-align:center; width:100%; margin-bottom:{margin_bottom}px;'>
            <div style='font-size:{title_size}px; font-weight:500; line-height:1.2; margin-bottom:6px;'>{title}</div>
            <div style='font-size:{value_size}px; font-weight:500; line-height:1.1;'>{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def add_total(df, col, total_value=None):
    if df.empty:
        return df.copy()
    if total_value is None:
        total_value = pd.to_numeric(df[col], errors='coerce').fillna(0).sum()
    total = pd.DataFrame({'Dia da Semana': ['TOTAL'], col: [int(total_value) if pd.notna(total_value) else 0]})
    return pd.concat([df, total], ignore_index=True)


def style_total(df):
    if df.empty:
        return df
    chave = None
    if 'Dia da Semana' in df.columns:
        chave = 'Dia da Semana'
    elif 'Ranking' in df.columns:
        chave = 'Ranking'
    if chave is None:
        return df
    return df.style.apply(
        lambda r: [('background-color:#EEF3F8;font-weight:700' if str(r.get(chave, '')).upper() == 'TOTAL' else '') for _ in r],
        axis=1,
    )


def ordenar_ano_semana(df, col):
    temp = df.copy()
    extraido = temp[col].astype(str).str.extract(r'(?P<ano>\d+)-S(?P<sem>\d+)')
    temp['_a'] = pd.to_numeric(extraido['ano'], errors='coerce').fillna(0)
    temp['_s'] = pd.to_numeric(extraido['sem'], errors='coerce').fillna(0)
    return temp.sort_values(['_a', '_s', col]).drop(columns=['_a', '_s'])


def annotate_bars(ax, bars, vals, formatter='pct', y_max=None):
    if y_max is None:
        y_max = ax.get_ylim()[1]
    for bar, val in zip(bars, vals):
        txt = fmt_pct(val) if formatter == 'pct' else fmt_num(val)

        if formatter == 'pct':
            # Somente os gráficos de indicadores ficam no modelo antigo:
            # rótulo interno quando couber; externo apenas quando truncaria.
            if float(val) <= 0:
                y = 1.2
                va = 'bottom'
            elif float(val) < 12:
                y = min(y_max - 1.0, float(val) + max(1.0, y_max * 0.02))
                va = 'bottom'
            else:
                y = max(1.3, float(val) * 0.53)
                va = 'center'
        else:
            # Rankings ficam no modelo novo.
            if float(val) >= (y_max * 0.12):
                y = float(val) * 0.55
                va = 'center'
            else:
                y = min(y_max - 1.0, float(val) + max(1.0, y_max * 0.02))
                va = 'bottom'

        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            txt,
            ha='center', va=va, rotation=90,
            fontsize=8, fontweight='bold', color='black', clip_on=False
        )


@st.cache_data(show_spinner=False)
def carregar(arquivo):
    df = pd.read_excel(arquivo, sheet_name=ABA, engine='openpyxl')
    falt = [c for c in REQ if c not in df.columns]
    if falt:
        raise ValueError('Colunas obrigatórias não encontradas: ' + ', '.join(falt))

    df = df.copy()
    df['Data Criação'] = pd.to_datetime(df['Data Criação'], errors='coerce')
    df['DT-EXP'] = pd.to_datetime(df['DT-EXP'], errors='coerce')
    df['dt_criacao_dia'] = df['Data Criação'].dt.normalize()
    df['dt_exp_dia'] = df['DT-EXP'].dt.normalize()

    for c in ['PEDIDO', 'PROTOCOLO', 'Cód Cliente']:
        df[c] = df[c].astype(str).str.strip()
        df.loc[df[c].isin(['nan', 'NaT', 'None']), c] = ''

    for c in ['CD Origem', 'LOJA (SAP)', 'Tipo Centro', 'UF_DEST']:
        df[c] = df[c].astype(str).str.strip().str.upper()

    df['REG_DEST'] = df['REG_DEST'].astype(str).str.strip()
    df['Custo Faturamento'] = pd.to_numeric(df['Custo Faturamento'], errors='coerce').fillna(0)
    if 'ROMANEIO' in df.columns:
        df['ROMANEIO'] = df['ROMANEIO'].astype(str).str.strip()
        df.loc[df['ROMANEIO'].isin(['nan', 'NaT', 'None']), 'ROMANEIO'] = ''
    else:
        df['ROMANEIO'] = ''

    df = df[(df['PEDIDO'] != '') & (df['CD Origem'] != '') & (df['LOJA (SAP)'] != '')].copy()

    ic = df['Data Criação'].dt.isocalendar()
    ie = df['DT-EXP'].dt.isocalendar()

    # Filtros por ano/mês usando ANO CALENDÁRIO
    df['Ano Filtro'] = df['Data Criação'].dt.year.astype('Int64')
    df['Semana Filtro'] = ic.week.astype('Int64')
    df['Mes Filtro'] = df['Data Criação'].dt.month.astype('Int64')
    df['Dia Filtro'] = df['Data Criação'].dt.day.astype('Int64')
    df['Ano-Semana Criação'] = ic.year.astype(str) + '-S' + df['Semana Filtro'].astype(str).str.zfill(2)
    df['Dia Semana Criação'] = df['Data Criação'].dt.dayofweek
    df['Dia Semana Criação Nome'] = df['Dia Semana Criação'].map(DIAS_MAP)

    df['Ano Filtro Exp'] = df['DT-EXP'].dt.year.astype('Int64')
    df['Semana Filtro Exp'] = ie.week.astype('Int64')
    df['Mes Filtro Exp'] = df['DT-EXP'].dt.month.astype('Int64')
    df['Dia Filtro Exp'] = df['DT-EXP'].dt.day.astype('Int64')
    df['Ano-Semana Exp'] = ie.year.astype(str) + '-S' + df['Semana Filtro Exp'].astype(str).str.zfill(2)
    df['Dia Semana Exp'] = df['DT-EXP'].dt.dayofweek
    df['Dia Semana Exp Nome'] = df['Dia Semana Exp'].map(DIAS_MAP)
    return df


def filtrar(df, anos, cds, lojas, meses, semanas, dias, ref='criacao'):
    out = df.copy()
    if cds:
        out = out[out['CD Origem'].isin(cds)]
    if lojas:
        out = out[out['LOJA (SAP)'].isin(lojas)]

    if ref == 'criacao':
        if anos:
            out = out[out['Ano Filtro'].isin(anos)]
        if meses:
            out = out[out['Mes Filtro'].isin(meses)]
        if semanas:
            out = out[out['Semana Filtro'].isin(semanas)]
        if dias:
            out = out[out['Dia Filtro'].isin(dias)]
    else:
        if anos:
            out = out[out['Ano Filtro Exp'].isin(anos)]
        if meses:
            out = out[out['Mes Filtro Exp'].isin(meses)]
        if semanas:
            out = out[out['Semana Filtro Exp'].isin(semanas)]
        if dias:
            out = out[out['Dia Filtro Exp'].isin(dias)]
    return out


def indicador1(df):
    if df.empty:
        return pd.DataFrame(columns=['Ano-Semana Criação', 'Indicador 1 (%)'])
    g = df[['Ano-Semana Criação', 'PEDIDO']].drop_duplicates().groupby('Ano-Semana Criação')['PEDIDO'].nunique().reset_index(name='q')
    g['Indicador 1 (%)'] = g['q'].apply(lambda x: score_limite(x, 6))
    return g[['Ano-Semana Criação', 'Indicador 1 (%)']]


def indicador2(df):
    if df.empty:
        return pd.DataFrame(columns=['Ano-Semana Criação', 'Indicador 2 (%)'])
    g = (
        df[['Ano-Semana Criação', 'PROTOCOLO']]
        .dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''")
        .drop_duplicates().groupby('Ano-Semana Criação')['PROTOCOLO'].nunique().reset_index(name='q')
    )
    g['Indicador 2 (%)'] = g['q'].apply(lambda x: score_limite(x, 2))
    return g[['Ano-Semana Criação', 'Indicador 2 (%)']]


def indicador3(df):
    if df.empty:
        return pd.DataFrame(columns=['Ano-Semana Exp', 'Indicador 3 (%)'])
    dia = (
        df[['Ano-Semana Exp', 'dt_exp_dia', 'PROTOCOLO']]
        .dropna(subset=['dt_exp_dia', 'PROTOCOLO']).query("PROTOCOLO != ''")
        .drop_duplicates().groupby(['Ano-Semana Exp', 'dt_exp_dia'])['PROTOCOLO'].nunique().reset_index(name='q')
    )
    dia['exc'] = (dia['q'] - 1).clip(lower=0)
    sem = dia.groupby('Ano-Semana Exp')['exc'].sum().reset_index(name='exc_total')
    sem['Indicador 3 (%)'] = sem['exc_total'].apply(score_excedente)
    return sem[['Ano-Semana Exp', 'Indicador 3 (%)']]


# ===== Séries semanais corretas para os gráficos =====
def indicador1_semanal_por_loja(base_cr):
    if base_cr.empty:
        return pd.DataFrame(columns=['Ano-Semana Criação', 'Indicador 1 (%)'])
    loja_semana = (
        base_cr[['CD Origem', 'LOJA (SAP)', 'Ano-Semana Criação', 'PEDIDO']]
        .drop_duplicates().groupby(['CD Origem', 'LOJA (SAP)', 'Ano-Semana Criação'])['PEDIDO']
        .nunique().reset_index(name='q')
    )
    loja_semana['Indicador 1 (%)'] = loja_semana['q'].apply(lambda x: score_limite(x, 6))
    semanal = loja_semana.groupby('Ano-Semana Criação')['Indicador 1 (%)'].mean().reset_index()
    semanal['Indicador 1 (%)'] = semanal['Indicador 1 (%)'].round(2)
    return semanal


def indicador2_semanal_por_loja(base_cr):
    if base_cr.empty:
        return pd.DataFrame(columns=['Ano-Semana Criação', 'Indicador 2 (%)'])
    loja_semana = (
        base_cr[['CD Origem', 'LOJA (SAP)', 'Ano-Semana Criação', 'PROTOCOLO']]
        .dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''")
        .drop_duplicates().groupby(['CD Origem', 'LOJA (SAP)', 'Ano-Semana Criação'])['PROTOCOLO']
        .nunique().reset_index(name='q')
    )
    loja_semana['Indicador 2 (%)'] = loja_semana['q'].apply(lambda x: score_limite(x, 2))
    semanal = loja_semana.groupby('Ano-Semana Criação')['Indicador 2 (%)'].mean().reset_index()
    semanal['Indicador 2 (%)'] = semanal['Indicador 2 (%)'].round(2)
    return semanal


def indicador3_semanal_por_loja(base_exp):
    if base_exp.empty:
        return pd.DataFrame(columns=['Ano-Semana Exp', 'Indicador 3 (%)'])
    dia_loja = (
        base_exp[['CD Origem', 'LOJA (SAP)', 'Ano-Semana Exp', 'dt_exp_dia', 'PROTOCOLO']]
        .dropna(subset=['dt_exp_dia', 'PROTOCOLO']).query("PROTOCOLO != ''")
        .drop_duplicates().groupby(['CD Origem', 'LOJA (SAP)', 'Ano-Semana Exp', 'dt_exp_dia'])['PROTOCOLO']
        .nunique().reset_index(name='q')
    )
    dia_loja['exc'] = (dia_loja['q'] - 1).clip(lower=0)
    semana_loja = dia_loja.groupby(['CD Origem', 'LOJA (SAP)', 'Ano-Semana Exp'])['exc'].sum().reset_index(name='exc_total')
    semana_loja['Indicador 3 (%)'] = semana_loja['exc_total'].apply(score_excedente)
    semanal = semana_loja.groupby('Ano-Semana Exp')['Indicador 3 (%)'].mean().reset_index()
    semanal['Indicador 3 (%)'] = semanal['Indicador 3 (%)'].round(2)
    return semanal


def resumo_semanais(base_cr, base_exp):
    bc = base_cr[base_cr['Dia Semana Criação'].isin(range(5))].copy()
    be = base_exp[base_exp['Dia Semana Exp'].isin(range(5))].copy()
    q1 = (
        bc[['PEDIDO', 'Dia Semana Criação Nome']].drop_duplicates()
        .groupby('Dia Semana Criação Nome')['PEDIDO'].nunique()
        .reindex(ORDEM_DIAS, fill_value=0).reset_index(name='Qtd Pedidos Criados')
        .rename(columns={'Dia Semana Criação Nome': 'Dia da Semana'})
    )
    q2 = (
        bc[['PROTOCOLO', 'Dia Semana Criação Nome']].dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''").drop_duplicates()
        .groupby('Dia Semana Criação Nome')['PROTOCOLO'].nunique()
        .reindex(ORDEM_DIAS, fill_value=0).reset_index(name='Qtd Protocolos')
        .rename(columns={'Dia Semana Criação Nome': 'Dia da Semana'})
    )
    q3 = (
        be[['PROTOCOLO', 'Dia Semana Exp Nome']].dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''").drop_duplicates()
        .groupby('Dia Semana Exp Nome')['PROTOCOLO'].nunique()
        .reindex(ORDEM_DIAS, fill_value=0).reset_index(name='Qtd Protocolos Expedidos')
        .rename(columns={'Dia Semana Exp Nome': 'Dia da Semana'})
    )
    total_q2 = bc.loc[bc['PROTOCOLO'].notna() & (bc['PROTOCOLO'] != ''), 'PROTOCOLO'].nunique()
    return add_total(q1, 'Qtd Pedidos Criados'), add_total(q2, 'Qtd Protocolos', total_q2), add_total(q3, 'Qtd Protocolos Expedidos')


def detalhe(df):
    if df.empty:
        return pd.DataFrame(columns=['Data Criação', 'Semana da OV', 'Data Expedição', 'CD Origem', 'Tipo Centro', 'REG_DEST', 'UF_DEST', 'PEDIDO', 'PROTOCOLO', 'ROMANEIO'])
    det = df[['Data Criação', 'Semana da OV', 'DT-EXP', 'CD Origem', 'Tipo Centro', 'REG_DEST', 'UF_DEST', 'PEDIDO', 'PROTOCOLO', 'ROMANEIO']].copy()
    det['Data Criação'] = det['Data Criação'].apply(data_extenso)
    det['Data Expedição'] = det['DT-EXP'].apply(data_extenso)
    det = det.drop(columns=['DT-EXP'])
    return det[['Data Criação', 'Semana da OV', 'Data Expedição', 'CD Origem', 'Tipo Centro', 'REG_DEST', 'UF_DEST', 'PEDIDO', 'PROTOCOLO', 'ROMANEIO']].sort_values(['Semana da OV', 'Data Criação', 'Data Expedição', 'CD Origem', 'PEDIDO']).reset_index(drop=True)


def tabela_consolidacao(df):
    if df.empty:
        return pd.DataFrame(columns=['Cód Cliente','LOJA (SAP)','CD Origem','Qtd Protocolos','Qtde de Pedidos','Custo por Pedido','Custo Faturamento'])
    ped = (
        df[['Cód Cliente','LOJA (SAP)','CD Origem','PEDIDO','Custo Faturamento']]
        .drop_duplicates(subset=['Cód Cliente','LOJA (SAP)','CD Origem','PEDIDO'])
        .groupby(['Cód Cliente','LOJA (SAP)','CD Origem'], dropna=False)
        .agg(**{'Qtde de Pedidos': ('PEDIDO', pd.Series.nunique), 'Custo por Pedido': ('Custo Faturamento', 'mean')})
        .reset_index()
    )
    ped['Custo Faturamento'] = ped['Qtde de Pedidos'] * ped['Custo por Pedido']
    prot = (
        df[['Cód Cliente','LOJA (SAP)','CD Origem','PROTOCOLO']]
        .dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''")
        .drop_duplicates(subset=['Cód Cliente','LOJA (SAP)','CD Origem','PROTOCOLO'])
        .groupby(['Cód Cliente','LOJA (SAP)','CD Origem'], dropna=False)['PROTOCOLO']
        .nunique().reset_index(name='Qtd Protocolos')
    )
    tb = ped.merge(prot, on=['Cód Cliente','LOJA (SAP)','CD Origem'], how='left')
    for c in ['Qtde de Pedidos','Qtd Protocolos','Custo por Pedido','Custo Faturamento']:
        tb[c] = pd.to_numeric(tb[c], errors='coerce').fillna(0)
    return tb[['Cód Cliente','LOJA (SAP)','CD Origem','Qtd Protocolos','Qtde de Pedidos','Custo por Pedido','Custo Faturamento']].sort_values(['Qtd Protocolos','Qtde de Pedidos','CD Origem','LOJA (SAP)'], ascending=[False,False,True,True]).reset_index(drop=True)


def formatar_consolidacao(df):
    out = df.copy()
    if out.empty:
        return out
    out['Custo por Pedido'] = out['Custo por Pedido'].apply(fmt_moeda)
    out['Custo Faturamento'] = out['Custo Faturamento'].apply(fmt_moeda)
    out['Qtd Protocolos'] = out['Qtd Protocolos'].astype(int)
    out['Qtde de Pedidos'] = out['Qtde de Pedidos'].astype(int)
    return out


def pivot_qtd(df, valor_col):
    if df.empty:
        return pd.DataFrame(columns=['Cód Cliente','LOJA (SAP)','Total'])
    temp = df[['Cód Cliente','LOJA (SAP)','Dia Filtro', valor_col]].dropna(subset=['Dia Filtro', valor_col]).copy()
    temp = temp[temp[valor_col].astype(str).str.strip() != '']
    temp = temp.drop_duplicates(subset=['Cód Cliente','LOJA (SAP)','Dia Filtro', valor_col])
    piv = temp.pivot_table(index=['Cód Cliente','LOJA (SAP)'], columns='Dia Filtro', values=valor_col, aggfunc='nunique', fill_value=0).reset_index()
    fix = ['Cód Cliente','LOJA (SAP)']
    dias = sorted([c for c in piv.columns if c not in fix], key=lambda x: int(x))
    piv['Total'] = piv[dias].sum(axis=1) if dias else 0
    return piv[fix + dias + ['Total']]


def pct_pedidos_tabela(df):
    cols = ['Cód Cliente','LOJA (SAP)','Mês','Semana','Qtd Pedidos','Indicador 1 (%)']
    if df.empty:
        return pd.DataFrame(columns=cols)

    pedidos = (
        df[['Cód Cliente','LOJA (SAP)','Mes Filtro','Semana Filtro','PEDIDO']]
        .dropna(subset=['Mes Filtro','Semana Filtro','PEDIDO'])
        .drop_duplicates()
    )
    grp = (
        pedidos.groupby(['Cód Cliente','LOJA (SAP)','Mes Filtro','Semana Filtro'])['PEDIDO']
        .nunique().reset_index(name='Qtd Pedidos')
    )

    grp['Indicador 1 (%)'] = grp['Qtd Pedidos'].apply(lambda x: score_limite(x, 6))
    grp = grp.rename(columns={'Mes Filtro':'Mês','Semana Filtro':'Semana'})
    grp['Mês Num'] = grp['Mês'].astype(int)
    grp['Semana'] = grp['Semana'].astype(int)
    grp['Mês'] = grp['Mês Num'].map(MESES_PT)
    grp = grp.sort_values(['Mês Num','Semana','Cód Cliente','LOJA (SAP)'])

    out = grp[['Cód Cliente','LOJA (SAP)','Mês','Semana','Qtd Pedidos','Indicador 1 (%)']].copy()
    out['Qtd Pedidos'] = out['Qtd Pedidos'].astype(int)
    out['Indicador 1 (%)'] = out['Indicador 1 (%)'].apply(fmt_pct)
    return out


def pct_protocolos_tabela(df):
    cols = ['Cód Cliente','LOJA (SAP)','Mês','Semana','Qtd Protocolos','Indicador 2 (%)']
    if df.empty:
        return pd.DataFrame(columns=cols)

    temp = (
        df[['Cód Cliente','LOJA (SAP)','Mes Filtro','Semana Filtro','PROTOCOLO']]
        .dropna(subset=['Mes Filtro','Semana Filtro','PROTOCOLO'])
        .query("PROTOCOLO != ''")
        .drop_duplicates()
    )
    grp = (
        temp.groupby(['Cód Cliente','LOJA (SAP)','Mes Filtro','Semana Filtro'])['PROTOCOLO']
        .nunique().reset_index(name='Qtd Protocolos')
    )
    grp['Indicador 2 (%)'] = grp['Qtd Protocolos'].apply(lambda x: score_limite(x, 2))
    grp = grp.rename(columns={'Mes Filtro':'Mês','Semana Filtro':'Semana'})
    grp['Mês Num'] = grp['Mês'].astype(int)
    grp['Semana'] = grp['Semana'].astype(int)
    grp['Mês'] = grp['Mês Num'].map(MESES_PT)
    grp = grp.sort_values(['Mês Num','Semana','Cód Cliente','LOJA (SAP)'])

    out = grp[['Cód Cliente','LOJA (SAP)','Mês','Semana','Qtd Protocolos','Indicador 2 (%)']].copy()
    out['Qtd Protocolos'] = out['Qtd Protocolos'].astype(int)
    out['Indicador 2 (%)'] = out['Indicador 2 (%)'].apply(fmt_pct)
    return out


def plot_rotulado(df, x_col, y_col, titulo):
    if df.empty:
        st.warning(f'Sem dados para {titulo}.')
        return
    temp = df.copy()
    if x_col in ['Ano-Semana Criação','Ano-Semana Exp']:
        temp = ordenar_ano_semana(temp, x_col)
    if MATPLOT:
        labels = temp[x_col].astype(str).tolist()
        vals = pd.to_numeric(temp[y_col], errors='coerce').fillna(0).tolist()
        largura = max(3.6, min(0.50 * len(labels) + 1.4, 5.4))
        fig, ax = plt.subplots(figsize=(largura, 2.9))
        width = 0.76 if len(labels) <= 4 else (0.68 if len(labels) <= 8 else 0.58)
        colors = ['#C62828' if float(v) == 0 else '#1565C0' for v in vals]
        bars = ax.bar(range(len(labels)), vals, color=colors, width=width)
        y_max = max(105, (max(vals) if vals else 0) * 1.20)
        ax.set_title(titulo, fontsize=9)
        ax.set_ylim(0, y_max)
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=90 if (len(labels) > 6 or x_col == 'Data') else 0, fontsize=7)
        ax.set_ylabel('%', fontsize=7)
        ax.tick_params(axis='y', labelsize=7)
        ax.grid(axis='y', alpha=0.3)
        annotate_bars(ax, bars, vals, formatter='pct', y_max=y_max)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.bar_chart(temp[[x_col,y_col]].set_index(x_col), use_container_width=True)
        txt = temp[[x_col,y_col]].copy()
        txt[y_col] = txt[y_col].apply(fmt_pct)
        st.dataframe(txt, hide_index=True, use_container_width=True)


def ranking_lojas(base_cr, base_exp):
    if base_cr.empty and base_exp.empty:
        return pd.DataFrame(columns=['Ranking','CD Origem','LOJA (SAP)','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação'])
    r1 = base_cr[['CD Origem','LOJA (SAP)','Ano-Semana Criação','PEDIDO']].drop_duplicates().groupby(['CD Origem','LOJA (SAP)','Ano-Semana Criação'])['PEDIDO'].nunique().reset_index(name='q1')
    r1['Indicador 1 (%)'] = r1['q1'].apply(lambda x: score_limite(x, 6))
    r2 = base_cr[['CD Origem','LOJA (SAP)','Ano-Semana Criação','PROTOCOLO']].dropna(subset=['PROTOCOLO']).query("PROTOCOLO != ''").drop_duplicates().groupby(['CD Origem','LOJA (SAP)','Ano-Semana Criação'])['PROTOCOLO'].nunique().reset_index(name='q2')
    r2['Indicador 2 (%)'] = r2['q2'].apply(lambda x: score_limite(x, 2))
    dia = base_exp[['CD Origem','LOJA (SAP)','Ano-Semana Exp','dt_exp_dia','PROTOCOLO']].dropna(subset=['dt_exp_dia','PROTOCOLO']).query("PROTOCOLO != ''").drop_duplicates().groupby(['CD Origem','LOJA (SAP)','Ano-Semana Exp','dt_exp_dia'])['PROTOCOLO'].nunique().reset_index(name='q3dia')
    dia['exc'] = (dia['q3dia'] - 1).clip(lower=0)
    r3 = dia.groupby(['CD Origem','LOJA (SAP)','Ano-Semana Exp'])['exc'].sum().reset_index(name='exc_total')
    r3['Indicador 3 (%)'] = r3['exc_total'].apply(score_excedente)
    s1 = r1.groupby(['CD Origem','LOJA (SAP)'])['Indicador 1 (%)'].mean().reset_index()
    s2 = r2.groupby(['CD Origem','LOJA (SAP)'])['Indicador 2 (%)'].mean().reset_index()
    s3 = r3.groupby(['CD Origem','LOJA (SAP)'])['Indicador 3 (%)'].mean().reset_index()
    rk = s1.merge(s2, on=['CD Origem','LOJA (SAP)'], how='outer').merge(s3, on=['CD Origem','LOJA (SAP)'], how='outer').fillna(0)
    rk['Score Ordenação'] = rk[['Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']].mean(axis=1).round(2)
    rk = rk.sort_values('Score Ordenação', ascending=True).reset_index(drop=True)
    rk.index = rk.index + 1
    rk['Ranking'] = rk.index.astype(str)
    return rk[['Ranking','CD Origem','LOJA (SAP)','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação']]


def totais_ranking_lojas(rk):
    if rk.empty:
        return {'Indicador 1 (%)': 0.0, 'Indicador 2 (%)': 0.0, 'Indicador 3 (%)': 0.0}
    return {
        'Indicador 1 (%)': round(float(pd.to_numeric(rk['Indicador 1 (%)'], errors='coerce').fillna(0).mean()), 2),
        'Indicador 2 (%)': round(float(pd.to_numeric(rk['Indicador 2 (%)'], errors='coerce').fillna(0).mean()), 2),
        'Indicador 3 (%)': round(float(pd.to_numeric(rk['Indicador 3 (%)'], errors='coerce').fillna(0).mean()), 2),
    }


def ranking_lojas_exibir(rk, incluir_total=True):
    cols = ['Ranking','CD Origem','LOJA (SAP)','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']
    if rk.empty:
        return pd.DataFrame(columns=cols)
    out = rk[cols].copy()
    if incluir_total:
        totais = totais_ranking_lojas(rk)
        total_row = pd.DataFrame([{'Ranking':'TOTAL','CD Origem':'','LOJA (SAP)':'','Indicador 1 (%)':totais['Indicador 1 (%)'],'Indicador 2 (%)':totais['Indicador 2 (%)'],'Indicador 3 (%)':totais['Indicador 3 (%)']}])
        out = pd.concat([out, total_row], ignore_index=True)
    for c in ['Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']:
        out[c] = out[c].apply(fmt_pct)
    return out


def plot_ranking_lojas(rk):
    if rk.empty:
        st.warning('Sem dados para Ranking das Lojas.')
        return
    temp = rk.head(TOP_LOJAS).copy()
    if MATPLOT:
        labels = (temp['CD Origem'].astype(str) + ' | ' + temp['LOJA (SAP)'].astype(str)).tolist()
        vals = pd.to_numeric(temp['Score Ordenação'], errors='coerce').fillna(0).tolist()
        n = len(temp)
        colors = plt.cm.RdYlGn(np.linspace(0,1,n))
        # Exatamente o mesmo tamanho do Ranking dos CDs
        fig, ax = plt.subplots(figsize=(max(10.0, min(0.72 * n + 6.0, 16.0)), 5.2))
        bars = ax.bar(range(n), vals, color=colors, width=0.78)
        y_max = max(100, (max(vals) if vals else 0) * 1.20)
        ax.set_title('Ranking das lojas - pior para melhor', fontsize=10)
        ax.set_ylim(0, y_max)
        ax.set_xticks(range(n))
        ax.set_xticklabels(labels, rotation=90, fontsize=8)
        ax.set_ylabel('Score', fontsize=9)
        ax.tick_params(axis='y', labelsize=8)
        ax.grid(axis='y', alpha=0.5)
        annotate_bars(ax, bars, vals, formatter='num', y_max=y_max)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        temp['Label'] = temp['CD Origem'].astype(str) + ' | ' + temp['LOJA (SAP)'].astype(str)
        st.bar_chart(temp[['Label','Score Ordenação']].set_index('Label'), use_container_width=True)


def ranking_cds(base_cr, base_exp):
    """
    Regra correta: cada linha do CD reflete o consolidado das lojas daquele CD.
    E o ranking é ordenado exatamente pela pior média consolidada das lojas.
    """
    rk_lojas = ranking_lojas(base_cr, base_exp)
    if rk_lojas.empty:
        return pd.DataFrame(columns=['Ranking','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação'])
    rk = (
        rk_lojas.groupby('CD Origem', dropna=False)
        .agg(**{
            'Qtd Lojas': ('LOJA (SAP)', 'nunique'),
            'Indicador 1 (%)': ('Indicador 1 (%)', 'mean'),
            'Indicador 2 (%)': ('Indicador 2 (%)', 'mean'),
            'Indicador 3 (%)': ('Indicador 3 (%)', 'mean'),
        })
        .reset_index()
    )
    rk['Score Ordenação'] = rk[['Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']].mean(axis=1).round(2)
    rk = rk.sort_values(['Score Ordenação','CD Origem'], ascending=[True, True]).reset_index(drop=True)
    rk.index = rk.index + 1
    rk['Ranking'] = rk.index.astype(str)
    return rk[['Ranking','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação']]


def totais_ranking_cds(rk):
    if rk.empty:
        return {'Qtd Lojas': 0, 'Indicador 1 (%)': 0.0, 'Indicador 2 (%)': 0.0, 'Indicador 3 (%)': 0.0, 'Score Ordenação': 0.0}
    return {
        'Qtd Lojas': int(pd.to_numeric(rk['Qtd Lojas'], errors='coerce').fillna(0).sum()),
        'Indicador 1 (%)': round(float(pd.to_numeric(rk['Indicador 1 (%)'], errors='coerce').fillna(0).mean()), 2),
        'Indicador 2 (%)': round(float(pd.to_numeric(rk['Indicador 2 (%)'], errors='coerce').fillna(0).mean()), 2),
        'Indicador 3 (%)': round(float(pd.to_numeric(rk['Indicador 3 (%)'], errors='coerce').fillna(0).mean()), 2),
        'Score Ordenação': round(float(pd.to_numeric(rk['Score Ordenação'], errors='coerce').fillna(0).mean()), 2),
    }


def ranking_cds_exibir(rk, incluir_total=True, totais_override=None):
    cols = ['Ranking','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score']
    if rk.empty:
        return pd.DataFrame(columns=cols)

    out = rk[['Ranking','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação']].copy()
    out = out.rename(columns={'Score Ordenação': 'Score'})

    if incluir_total:
        totais = totais_override if totais_override is not None else totais_ranking_cds(rk)
        total_row = pd.DataFrame([{
            'Ranking':'TOTAL',
            'CD Origem':'',
            'Qtd Lojas': totais.get('Qtd Lojas', ''),
            'Indicador 1 (%)': totais['Indicador 1 (%)'],
            'Indicador 2 (%)': totais['Indicador 2 (%)'],
            'Indicador 3 (%)': totais['Indicador 3 (%)'],
            'Score': totais.get('Score Ordenação', 0.0)
        }])
        out = pd.concat([out, total_row], ignore_index=True)

    for c in ['Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']:
        out[c] = out[c].apply(fmt_pct)
    out['Score'] = out['Score'].apply(fmt_num)
    out['Qtd Lojas'] = out['Qtd Lojas'].apply(lambda x: '' if x == '' else int(float(x)) if pd.notna(x) else '')
    return out


def plot_ranking_cds(rk):
    if rk.empty:
        st.warning('Sem dados para Ranking dos CDs.')
        return
    temp = rk.head(TOP_LOJAS).copy()
    if MATPLOT:
        labels = temp['CD Origem'].astype(str).tolist()
        vals = pd.to_numeric(temp['Score Ordenação'], errors='coerce').fillna(0).tolist()
        n = len(temp)
        colors = plt.cm.RdYlGn(np.linspace(0,1,n))
        # Exatamente o mesmo tamanho do Ranking das lojas
        fig, ax = plt.subplots(figsize=(max(10.0, min(0.72 * n + 6.0, 16.0)), 5.2))
        bars = ax.bar(range(n), vals, color=colors, width=0.78)
        y_max = max(100, (max(vals) if vals else 0) * 1.20)
        ax.set_title('Ranking dos CDs - pior para melhor', fontsize=10)
        ax.set_ylim(0, y_max)
        ax.set_xticks(range(n))
        ax.set_xticklabels(labels, rotation=90, fontsize=8)
        ax.set_ylabel('Score', fontsize=10)
        ax.tick_params(axis='y', labelsize=8)
        ax.grid(axis='y', alpha=0.3)
        annotate_bars(ax, bars, vals, formatter='num', y_max=y_max)
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        st.bar_chart(temp[['CD Origem','Score Ordenação']].set_index('CD Origem'), use_container_width=True)


def ranking_cds_comparativo_mes_a_mes(base_cr, base_exp):
    cols = ['Ranking','Mês','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score']
    if base_cr.empty and base_exp.empty:
        return pd.DataFrame(columns=cols)

    meses_cr = set(pd.to_numeric(base_cr['Mes Filtro'], errors='coerce').dropna().astype(int).tolist()) if (not base_cr.empty and 'Mes Filtro' in base_cr.columns) else set()
    meses_exp = set(pd.to_numeric(base_exp['Mes Filtro Exp'], errors='coerce').dropna().astype(int).tolist()) if (not base_exp.empty and 'Mes Filtro Exp' in base_exp.columns) else set()
    meses = sorted(meses_cr.union(meses_exp))

    blocos = []
    for mes in meses:
        cr_mes = base_cr[base_cr['Mes Filtro'] == mes].copy() if (not base_cr.empty and 'Mes Filtro' in base_cr.columns) else base_cr.copy()
        exp_mes = base_exp[base_exp['Mes Filtro Exp'] == mes].copy() if (not base_exp.empty and 'Mes Filtro Exp' in base_exp.columns) else base_exp.copy()
        rk_mes = ranking_cds(cr_mes, exp_mes)
        if rk_mes.empty:
            continue

        bloco = rk_mes[['Ranking','CD Origem','Qtd Lojas','Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)','Score Ordenação']].copy()
        bloco.insert(1, 'Mês', MESES_PT.get(int(mes), str(mes)))
        bloco = bloco.rename(columns={'Score Ordenação': 'Score'})

        total_row = pd.DataFrame([{
            'Ranking': 'TOTAL',
            'Mês': MESES_PT.get(int(mes), str(mes)),
            'CD Origem': '',
            'Qtd Lojas': int(pd.to_numeric(rk_mes['Qtd Lojas'], errors='coerce').fillna(0).sum()),
            'Indicador 1 (%)': round(float(pd.to_numeric(rk_mes['Indicador 1 (%)'], errors='coerce').fillna(0).mean()), 2),
            'Indicador 2 (%)': round(float(pd.to_numeric(rk_mes['Indicador 2 (%)'], errors='coerce').fillna(0).mean()), 2),
            'Indicador 3 (%)': round(float(pd.to_numeric(rk_mes['Indicador 3 (%)'], errors='coerce').fillna(0).mean()), 2),
            'Score': round(float(pd.to_numeric(rk_mes['Score Ordenação'], errors='coerce').fillna(0).mean()), 2),
        }])
        bloco = pd.concat([bloco, total_row], ignore_index=True)
        blocos.append(bloco)

    if not blocos:
        return pd.DataFrame(columns=cols)

    out = pd.concat(blocos, ignore_index=True)
    for c in ['Indicador 1 (%)','Indicador 2 (%)','Indicador 3 (%)']:
        out[c] = out[c].apply(fmt_pct)
    out['Score'] = out['Score'].apply(fmt_num)
    out['Qtd Lojas'] = out['Qtd Lojas'].apply(lambda x: '' if x == '' else int(float(x)) if pd.notna(x) else '')
    return out[cols]

st.title('Indicadores dos Pedidos para LPs')
st.caption('Revisão aplicada: protocolos distintos com contagem única no total. Nova regra dos indicadores: se ultrapassar o limite do indicador, o percentual é zerado.')

with st.sidebar:
    st.header('Origem dos dados')
    arquivo = st.file_uploader('Selecione o arquivo Excel', type=['xlsx'])
    arquivo_padrao_existe = os.path.exists(ARQUIVO_PADRAO)
    usar_padrao = st.checkbox('Usar arquivo padrão da pasta do projeto', value=arquivo_padrao_existe, disabled=not arquivo_padrao_existe)
    if not arquivo_padrao_existe:
        st.caption(f'Arquivo padrão não encontrado na pasta: {ARQUIVO_PADRAO}')

try:
    if arquivo is not None:
        df = carregar(arquivo)
    elif usar_padrao and os.path.exists(ARQUIVO_PADRAO):
        df = carregar(ARQUIVO_PADRAO)
    else:
        st.warning('Faça upload do Excel para continuar. O arquivo padrão não foi encontrado na pasta do projeto.')
        st.stop()
except Exception as e:
    st.error(f'Erro ao carregar a base: {e}')
    st.stop()

with st.sidebar:
    st.header('Filtros')
    anos = sorted([int(x) for x in df['Ano Filtro'].dropna().unique().tolist()])
    ano_atual = datetime.now().year
    ano_default = [ano_atual] if ano_atual in anos else ([max(anos)] if anos else [])
    sel_anos = st.multiselect('Ano', anos, default=ano_default)
    cds_disp = sorted(df[df['Ano Filtro'].isin(sel_anos)]['CD Origem'].dropna().unique().tolist()) if sel_anos else sorted(df['CD Origem'].dropna().unique().tolist())
    sel_cds = st.multiselect('CD Origem', cds_disp)
    lojas_base = df.copy()
    if sel_anos:
        lojas_base = lojas_base[lojas_base['Ano Filtro'].isin(sel_anos)]
    if sel_cds:
        lojas_base = lojas_base[lojas_base['CD Origem'].isin(sel_cds)]
    lojas_disp = sorted(lojas_base['LOJA (SAP)'].dropna().unique().tolist())
    sel_lojas = st.multiselect('Loja', lojas_disp)
    meses_base = df.copy()
    if sel_anos:
        meses_base = meses_base[meses_base['Ano Filtro'].isin(sel_anos)]
    if sel_cds:
        meses_base = meses_base[meses_base['CD Origem'].isin(sel_cds)]
    if sel_lojas:
        meses_base = meses_base[meses_base['LOJA (SAP)'].isin(sel_lojas)]
    meses_nums = sorted([int(x) for x in meses_base['Mes Filtro'].dropna().unique().tolist()])
    meses_options = {MESES_PT[m]: m for m in meses_nums}
    sel_meses_nomes = st.multiselect('Mês', list(meses_options.keys()))
    sel_meses = [meses_options[m] for m in sel_meses_nomes]
    semanas_base = meses_base.copy()
    if sel_meses:
        semanas_base = semanas_base[semanas_base['Mes Filtro'].isin(sel_meses)]
    semanas_disp = sorted([int(x) for x in semanas_base['Semana Filtro'].dropna().unique().tolist()])
    sel_semanas = st.multiselect('Semana', semanas_disp)
    dias_base = semanas_base.copy()
    if sel_semanas:
        dias_base = dias_base[dias_base['Semana Filtro'].isin(sel_semanas)]
    dias_disp = sorted([int(x) for x in dias_base['Dia Filtro'].dropna().unique().tolist()])
    sel_dias = st.multiselect('Dia', dias_disp)

base_cr = filtrar(df, sel_anos, sel_cds, sel_lojas, sel_meses, sel_semanas, sel_dias, 'criacao')
base_exp = filtrar(df, sel_anos, sel_cds, sel_lojas, sel_meses, sel_semanas, sel_dias, 'exp')

# totais gerais alinhados ao ranking das lojas
rk = ranking_lojas(base_cr, base_exp)
totais_rk_lojas = totais_ranking_lojas(rk)

# séries corretas dos gráficos (média semanal consolidada das lojas)
i1 = indicador1_semanal_por_loja(base_cr)
i2 = indicador2_semanal_por_loja(base_cr)
i3 = indicador3_semanal_por_loja(base_exp)

v1 = fmt_pct(totais_rk_lojas['Indicador 1 (%)'])
v2 = fmt_pct(totais_rk_lojas['Indicador 2 (%)'])
v3 = fmt_pct(totais_rk_lojas['Indicador 3 (%)'])

c1, c2, c3 = st.columns(3)
with c1:
    render_metric_centered('Indicador 1 - Pedidos criados por semana', v1, title_size=15, value_size=40, margin_bottom=2)
with c2:
    render_metric_centered('Indicador 2 - Protocolos por dia na semana', v2, title_size=15, value_size=40, margin_bottom=2)
with c3:
    render_metric_centered('Indicador 3 - Expedições por dia', v3, title_size=15, value_size=40, margin_bottom=2)

st.subheader('Quantidade por dia da semana (SEG a SEX)')
q1, q2, q3 = resumo_semanais(base_cr, base_exp)
cc1, cc2, cc3 = st.columns(3)
with cc1:
    st.markdown('**Indicador 1 - quantidade de pedidos criados**')
    st.dataframe(style_total(q1), hide_index=True, use_container_width=True)
with cc2:
    st.markdown('**Indicador 2 - quantidade de protocolos por dia na semana**')
    st.dataframe(style_total(q2), hide_index=True, use_container_width=True)
with cc3:
    st.markdown('**Indicador 3 - quantidade de protocolos expedidos por dia**')
    st.dataframe(style_total(q3), hide_index=True, use_container_width=True)

st.subheader('Detalhamento de pedidos')
det = detalhe(base_cr)
st.dataframe(det, hide_index=True, use_container_width=True, height=360)
det_buf = BytesIO()
with pd.ExcelWriter(det_buf, engine='openpyxl') as writer:
    det.to_excel(writer, sheet_name='Detalhamento', index=False)
st.download_button('Baixar detalhamento em Excel', data=det_buf.getvalue(), file_name='detalhamento_pedidos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.divider()
st.subheader('Indicadores gráficos')
g1, g2, g3 = st.columns(3)
with g1:
    render_metric_centered('Média geral do Indicador 1', v1, title_size=15, value_size=40, margin_bottom=8)
    plot_rotulado(i1, 'Ano-Semana Criação', 'Indicador 1 (%)', 'Indicador 1 - quantidade de pedidos criados')
with g2:
    render_metric_centered('Média geral do Indicador 2', v2, title_size=15, value_size=40, margin_bottom=8)
    plot_rotulado(i2, 'Ano-Semana Criação', 'Indicador 2 (%)', 'Indicador 2 - quantidade de protocolos por dia na semana')
with g3:
    render_metric_centered('Média geral do Indicador 3', v3, title_size=15, value_size=40, margin_bottom=8)
    plot_rotulado(i3, 'Ano-Semana Exp', 'Indicador 3 (%)', 'Indicador 3 - quantidade de protocolos expedidos por dia')

st.subheader('Ranking das lojas')
rk_exibir = ranking_lojas_exibir(rk, incluir_total=True)
st.dataframe(style_total(rk_exibir), hide_index=True, use_container_width=True, height=420)
plot_ranking_lojas(rk)

st.subheader('Ranking dos CDs')
rk_cd = ranking_cds(base_cr, base_exp)
rk_cd_exibir = ranking_cds_exibir(
    rk_cd,
    incluir_total=True,
    totais_override={
        'Qtd Lojas': int(rk['LOJA (SAP)'].nunique()) if not rk.empty else 0,
        **totais_rk_lojas,
        'Score Ordenação': round(float(pd.to_numeric(rk['Score Ordenação'], errors='coerce').fillna(0).mean()), 2) if not rk.empty else 0.0
    }
)
st.dataframe(style_total(rk_cd_exibir), hide_index=True, use_container_width=True, height=420)
plot_ranking_cds(rk_cd)

st.subheader('Ranking dos CDs - Comparativo - Mês a Mês')
rk_cd_mes_a_mes = ranking_cds_comparativo_mes_a_mes(base_cr, base_exp)
st.dataframe(style_total(rk_cd_mes_a_mes), hide_index=True, use_container_width=True, height=520)

st.subheader('Tabela analítica de consolidação por loja')
st.dataframe(formatar_consolidacao(tabela_consolidacao(base_cr)), hide_index=True, use_container_width=True, height=420)

st.subheader('Quantidade de Pedidos por Dia - Semana - Mês')
st.dataframe(pivot_qtd(base_cr, 'PEDIDO'), hide_index=True, use_container_width=True, height=340)

st.subheader('Quantidade de Protocolos distintos por Dia - Semana - Mês')
st.dataframe(pivot_qtd(base_cr, 'PROTOCOLO'), hide_index=True, use_container_width=True, height=340)

st.subheader('Percentual de Pedidos por Dia - Semana - Mês')
st.dataframe(pct_pedidos_tabela(base_cr), hide_index=True, use_container_width=True, height=340)

st.subheader('Percentual de Protocolos por Dia - Semana - Mês')
st.dataframe(pct_protocolos_tabela(base_cr), hide_index=True, use_container_width=True, height=340)

buf = BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    det.to_excel(writer, sheet_name='Detalhamento', index=False)
    q1.to_excel(writer, sheet_name='Resumo_Ind1_Semana', index=False)
    q2.to_excel(writer, sheet_name='Resumo_Ind2_Semana', index=False)
    q3.to_excel(writer, sheet_name='Resumo_Ind3_Semana', index=False)
    i1.to_excel(writer, sheet_name='Indicador_1', index=False)
    i2.to_excel(writer, sheet_name='Indicador_2', index=False)
    i3.to_excel(writer, sheet_name='Indicador_3', index=False)
    rk_exibir.to_excel(writer, sheet_name='Ranking_Lojas', index=False)
    rk_cd_exibir.to_excel(writer, sheet_name='Ranking_CDs', index=False)
    rk_cd_mes_a_mes.to_excel(writer, sheet_name='Ranking_CDs_Mes_a_Mes', index=False)
    pct_pedidos_tabela(base_cr).to_excel(writer, sheet_name='Pct_Pedidos', index=False)
    pct_protocolos_tabela(base_cr).to_excel(writer, sheet_name='Pct_Protocolos', index=False)
    tabela_consolidacao(base_cr).to_excel(writer, sheet_name='Tabela_Consolidacao', index=False)

st.download_button('Baixar resultado em Excel', data=buf.getvalue(), file_name='resultado_indicadores_lps.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
