import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime

# Configuração da página original (layout wide padrão)
st.set_page_config(page_title="Portal de Relatórios Operacionais", layout="wide")

# O ID do seu ficheiro do Google Drive
FILE_ID = "1T2HZveStvaxx3TByMYB6zCjaao1v9gh8"
url_drive = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"

# --- ESTILIZAÇÃO CSS AVANÇADA ---
st.markdown("""
    <style>
    /* 1. LETRA DAS GUIAS NO TOPO */
    div[data-testid="stPills"] button {
        font-size: 15px !important;
        font-weight: 600 !important;
        padding: 6px 14px !important;
    }
    
    /* 2. CARDS DE MÉTRICAS */
    .card-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; }
    .kpi-card { background-color: #ffffff; padding: 10px 18px; border-radius: 8px; border: 1px solid #e1e4e8; box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.04); flex: 1; font-family: sans-serif; }
    .card-blue { border-top: 4px solid #2b7bba; }
    .card-purple { border-top: 4px solid #7b3294; }
    .card-title { color: #586069; font-size: 13px; font-weight: 600; margin-bottom: 2px; }
    .card-value { color: #1f2328; font-size: 24px; font-weight: 700; margin-bottom: 2px; }
    .card-sub { color: #657180; font-size: 11px; }
    
    /* 3. ALINHAMENTO GERAL DOS CONTAINERS */
    .chart-card { background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1px solid #e1e4e8; box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.04); margin-bottom: 25px; width: 100%; }
    .chart-card-compact { background-color: #ffffff; padding: 8px; border-radius: 6px; border: 1px solid #e1e4e8; margin-bottom: 10px; }
    
    /* 4. TABELAS CUSTOMIZADAS EM HTML - Largura 85% para caber o TME com folga */
    .custom-table { width: 85%; margin-left: auto; margin-right: auto; border-collapse: collapse; font-family: sans-serif; font-size: 14px; margin-top: 10px; }
    .custom-table th { background-color: #1f497d; color: #ffffff; font-weight: bold; text-align: center; padding: 8px; border: 1px solid #e1e4e8; }
    .custom-table td { border: 1px solid #e1e4e8; padding: 6px; text-align: center; }
    .td-dia { font-weight: bold; background-color: #ffffff; color: #000000; }
    .td-vazio { background-color: #ffffff; color: #000000; }
    .td-vermelho { background-color: #fce4e4; color: #c00000; font-weight: bold; }
    .td-verde { background-color: #e2efda; color: #38761d; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE FORMATADORES E TABELAS HTML ---
def formatar_para_hhmmss(minutos):
    if pd.isna(minutos) or minutos < 0: return "00:00:00"
    total_segundos = int(round(minutos * 60))
    return f"{total_segundos // 3600:02d}:{(total_segundos % 3600) // 60:02d}:{total_segundos % 60:02d}"

def formatar_para_ms(minutos):
    if pd.isna(minutos) or minutos < 0: return "0m00s"
    total_segundos = int(round(minutos * 60))
    return f"{total_segundos // 60}m{total_segundos % 60:02d}s"

def gerar_celula_html(val):
    if pd.isna(val) or val == "": return "<td class='td-vazio'>—</td>"
    if val < 0.90: return f"<td class='td-vermelho'>{val*100:.1f}%</td>"
    return f"<td class='td-verde'>{val*100:.1f}%</td>"

def renderizar_tabela_html(df_pivot):
    html = "<table class='custom-table'><tr>"
    for col in df_pivot.columns: html += f"<th>{col}</th>"
    html += "</tr>"
    for _, row in df_pivot.iterrows():
        html += "<tr>"
        html += f"<td class='td-dia'>{row['Dia']}</td>"
        for col in df_pivot.columns[1:]: html += gerar_celula_html(row[col])
        html += "</tr>"
    html += "</table>"
    return html

def renderizar_tabela_ns_tma_html(df):
    html = "<table class='custom-table'><tr>"
    for col in df.columns: html += f"<th>{col}</th>"
    html += "</tr>"
    for _, row in df.iterrows():
        if row['Operação'] == 'GERAL': html += "<tr style='background-color: #f1f3f4; font-weight: bold;'>"
        else: html += "<tr>"
        for col in df.columns: html += f"<td>{row[col]}</td>"
        html += "</tr>"
    html += "</table>"
    return html

def renderizar_tabela_diario_rep_html(df):
    html = "<table class='custom-table'><tr>"
    for col in df.columns: html += f"<th>{col}</th>"
    html += "</tr>"
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = str(row[col])
            if col in ['NS Msg', 'NS Voz'] and '%' in val:
                try:
                    v = float(val.replace('%', ''))
                    if v < 90.0: html += f"<td class='td-vermelho'>{val}</td>"
                    else: html += f"<td class='td-verde'>{val}</td>"
                except: html += f"<td>{val}</td>"
            else: html += f"<td>{val}</td>"
        html += "</tr>"
    html += "</table>"
    return html

# --- EXTRATOR INTELIGENTE DE DADOS ---
@st.cache_data(ttl=600)
def carregar_todos_os_dados(url):
    df_main = pd.read_excel(url, header=3).dropna(subset=['Operação'])
    
    # Adicionando a leitura inteligente do TME também
    colunas_tempo = ['TMA Mensageria', 'TME Mensageria', 'TMA Telefonia', 'TME', 'TME Telefonia']
    for col in colunas_tempo:
        if col in df_main.columns:
            df_main[col] = pd.to_timedelta(df_main[col].astype(str), errors='coerce')
            df_main[f'{col} (Minutos)'] = df_main[col].dt.total_seconds() / 60
            
    df_raw = pd.read_excel(url, sheet_name="Diário por Operação", header=None)
    
    ops_conhecidas = ['SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
    linhas_processadas = []
    op_atual = None
    
    for index, row in df_raw.iterrows():
        val_a = row[0]
        if pd.isna(val_a):
            continue
            
        str_a = str(val_a).strip().upper()
        if str_a in ops_conhecidas:
            op_atual = str_a
            continue
            
        if op_atual:
            dia_limpo = None
            
            if isinstance(val_a, pd.Timestamp) or isinstance(val_a, datetime):
                dia_limpo = val_a.strftime("%d/%m")
            elif str_a[0].isdigit() and ('/' in str_a or '-' in str_a):
                texto_data = str_a.split(' ')[0] 
                if '-' in texto_data and len(texto_data.split('-')) == 3: 
                    partes = texto_data.split('-')
                    dia_limpo = f"{partes[2]}/{partes[1]}"
                else:
                    dia_limpo = texto_data.replace('-', '/')
                    
            if dia_limpo:
                def limpa_vol(v):
                    if pd.isna(v) or str(v).strip() == '': return 0
                    try: return int(str(v).replace('.', '').replace(',', ''))
                    except: return 0
                    
                def limpa_ns(v):
                    if pd.isna(v) or str(v).strip() == '': return np.nan
                    if isinstance(v, (int, float)): return float(v)
                    try: return float(str(v).replace('%', '').replace(',', '.')) / 100
                    except: return np.nan
                    
                def limpa_tma(v):
                    if pd.isna(v) or str(v).strip() == '': return 0.0
                    try:
                        if hasattr(v, 'hour'): return v.hour * 60 + v.minute + v.second / 60
                        partes = str(v).split(':')
                        return int(partes[0]) * 60 + int(partes[1]) + int(partes[2]) / 60
                    except: return 0.0

                linhas_processadas.append({
                    "Dia": dia_limpo,
                    "Operação": op_atual,
                    "Vol Msg": limpa_vol(row[1]),
                    "NS Msg": limpa_ns(row[2]),
                    "TMA Msg (Min)": limpa_tma(row[3]),
                    "Vol Voz": limpa_vol(row[4]),
                    "NS Voz": limpa_ns(row[5]),
                    "TMA Voz (Min)": limpa_tma(row[6])
                })
    
    cols_oficiais = ["Dia", "Operação", "Vol Msg", "NS Msg", "TMA Msg (Min)", "Vol Voz", "NS Voz", "TMA Voz (Min)"]
    df_diario = pd.DataFrame(linhas_processadas, columns=cols_oficiais)
    
    geral_rows = []
    if not df_diario.empty:
        for dia in df_diario['Dia'].unique():
            df_dia = df_diario[df_diario['Dia'] == dia]
            v_msg = df_dia['Vol Msg'].sum()
            v_voz = df_dia['Vol Voz'].sum()
            
            ns_m = (df_dia['NS Msg'] * df_dia['Vol Msg']).sum() / v_msg if v_msg > 0 else np.nan
            ns_v = (df_dia['NS Voz'] * df_dia['Vol Voz']).sum() / v_voz if v_voz > 0 else np.nan
            tma_m = (df_dia['TMA Msg (Min)'] * df_dia['Vol Msg']).sum() / v_msg if v_msg > 0 else 0
            tma_v = (df_dia['TMA Voz (Min)'] * df_dia['Vol Voz']).sum() / v_voz if v_voz > 0 else 0
            
            geral_rows.append({"Dia": dia, "Operação": "GERAL", "Vol Msg": v_msg, "NS Msg": ns_m, "TMA Msg (Min)": tma_m, "Vol Voz": v_voz, "NS Voz": ns_v, "TMA Voz (Min)": tma_v})
        
        df_diario = pd.concat([df_diario, pd.DataFrame(geral_rows)], ignore_index=True)
    
    return df_main, df_diario

try: 
    df_resumo, df_diario = carregar_todos_os_dados(url_drive)
except Exception as e: 
    st.error(f"Erro ao conectar com a base. Verifique os dados no Excel: {e}")
    st.stop()

# --- MENU PILLS ---
col_menu, col_btn = st.columns([8, 2])
with col_menu:
    aba_selecionada = st.pills(label="Menu Principal", options=["Geral", "NS por Operação", "NS Diário (perdas)", "Tabela Consolidada", "Diário por Operação"], default="Geral", label_visibility="collapsed")
with col_btn:
    if st.button("🔄 Atualizar Dados", use_container_width=True): st.cache_data.clear(); st.rerun()

st.markdown("---")

# ==============================================================================
# 1. GUIA: GERAL
# ==============================================================================
if aba_selecionada == "Geral":
    df_geral = df_resumo[df_resumo['Operação'] == 'GERAL']
    df_ops = df_resumo[df_resumo['Operação'] != 'GERAL']
    
    ns_msg = f"{df_geral['NS Mensageria'].values[0] * 100:.1f}%" if not df_geral.empty else "60.8%"
    ns_voz = f"{df_geral['NS Telefonia'].values[0] * 100:.1f}%" if not df_geral.empty else "64.6%"
    tma_msg_str = formatar_para_ms(df_geral['TMA Mensageria (Minutos)'].values[0]) if not df_geral.empty else "37m56s"
    tma_voz_str = formatar_para_ms(df_geral['TMA Telefonia (Minutos)'].values[0]) if not df_geral.empty else "6m56s"

    # Verificando as colunas de TME
    tme_msg_str = formatar_para_ms(df_geral['TME Mensageria (Minutos)'].values[0]) if 'TME Mensageria (Minutos)' in df_geral.columns and not df_geral.empty else "—"
    
    if 'TME Telefonia (Minutos)' in df_geral.columns:
        tme_voz_str = formatar_para_ms(df_geral['TME Telefonia (Minutos)'].values[0]) if not df_geral.empty else "—"
    elif 'TME (Minutos)' in df_geral.columns:
        tme_voz_str = formatar_para_ms(df_geral['TME (Minutos)'].values[0]) if not df_geral.empty else "—"
    else:
        tme_voz_str = "—"

    # Calculando os volumes TOTAIS reais a partir da aba Diário
    tot_msg_geral = df_diario[df_diario['Operação'] != 'GERAL']['Vol Msg'].sum()
    tot_voz_geral = df_diario[df_diario['Operação'] != 'GERAL']['Vol Voz'].sum()
    vol_msg_str = f"{tot_msg_geral:,.0f}".replace(',', '.') + " atend."
    vol_voz_str = f"{tot_voz_geral:,.0f}".replace(',', '.') + " chamadas"

    st.title("Visão Geral")
    
    # Nova estrutura de cards: Linha 1 (Msg), Linha 2 (Voz)
    html_cards = f"""
    <div class="card-container">
        <div class="kpi-card card-blue"><div class="card-title">NS Mensageria</div><div class="card-value">{ns_msg}</div><div class="card-sub">{vol_msg_str}</div></div>
        <div class="kpi-card card-blue"><div class="card-title">TMA Mensageria</div><div class="card-value">{tma_msg_str}</div><div class="card-sub">meta 90% &le; 4min</div></div>
        <div class="kpi-card card-blue"><div class="card-title">TME Mensageria</div><div class="card-value">{tme_msg_str}</div><div class="card-sub">tempo médio de espera</div></div>
    </div>
    <div class="card-container">
        <div class="kpi-card card-purple"><div class="card-title">NS Telefonia</div><div class="card-value">{ns_voz}</div><div class="card-sub">{vol_voz_str}</div></div>
        <div class="kpi-card card-purple"><div class="card-title">TMA Telefonia</div><div class="card-value">{tma_voz_str}</div><div class="card-sub">meta 90% &le; 10s</div></div>
        <div class="kpi-card card-purple"><div class="card-title">TME Telefonia</div><div class="card-value">{tme_voz_str}</div><div class="card-sub">tempo médio de espera</div></div>
    </div>
    """
    st.markdown(html_cards, unsafe_allow_html=True)

    # Gráficos agora em 3 colunas para incluir o TME
    colg1, colg2, colg3 = st.columns(3)
    with colg1:
        st.subheader("🎯 Nível de Serviço (NS)")
        df_melt = df_ops.melt(id_vars=['Operação'], value_vars=['NS Mensageria', 'NS Telefonia'], var_name='Canal', value_name='Val')
        df_melt['Porcentagem (%)'] = df_melt['Val'] * 100
        df_melt['Canal'] = df_melt['Canal'].map({'NS Mensageria': 'NS Msg', 'NS Telefonia': 'NS Voz'})
        fig = px.bar(df_melt, x='Operação', y='Porcentagem (%)', color='Canal', barmode='group', text='Porcentagem (%)', color_discrete_sequence=['#2b7bba', '#7b3294'], height=290)
        fig.update_yaxes(ticksuffix="%", range=[0, 100])
        fig.update_layout(xaxis_title="", yaxis_title="", showlegend=True, margin=dict(l=30, r=10, t=10, b=30), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"))
        fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with colg2:
        st.subheader("⏱️ Tempo de Atendimento (TMA)")
        df_melt_t = df_ops.melt(id_vars=['Operação'], value_vars=['TMA Mensageria (Minutos)', 'TMA Telefonia (Minutos)'], var_name='Canal', value_name='Min')
        df_melt_t['Canal'] = df_melt_t['Canal'].map({'TMA Mensageria (Minutos)': 'TMA Msg', 'TMA Telefonia (Minutos)': 'TMA Voz'})
        df_melt_t['Texto_hhmmss'] = df_melt_t['Min'].apply(formatar_para_hhmmss)
        fig2 = px.bar(df_melt_t, x='Operação', y='Min', color='Canal', barmode='group', text='Texto_hhmmss', color_discrete_sequence=['#2b7bba', '#7b3294'], height=290)
        fig2.update_layout(xaxis_title="", yaxis_title="minutos", showlegend=True, margin=dict(l=30, r=10, t=10, b=30), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"))
        fig2.update_traces(textposition='outside')
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig2, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with colg3:
        st.subheader("⏳ Tempo de Espera (TME)")
        tme_voz_col = 'TME Telefonia (Minutos)' if 'TME Telefonia (Minutos)' in df_ops.columns else ('TME (Minutos)' if 'TME (Minutos)' in df_ops.columns else None)
        tme_msg_col = 'TME Mensageria (Minutos)' if 'TME Mensageria (Minutos)' in df_ops.columns else None
        
        if tme_msg_col and tme_voz_col:
            df_melt_tme = df_ops.melt(id_vars=['Operação'], value_vars=[tme_msg_col, tme_voz_col], var_name='Canal', value_name='Min')
            df_melt_tme['Canal'] = df_melt_tme['Canal'].map({tme_msg_col: 'TME Msg', tme_voz_col: 'TME Voz'})
            df_melt_tme['Texto_hhmmss'] = df_melt_tme['Min'].apply(formatar_para_hhmmss)
            fig3 = px.bar(df_melt_tme, x='Operação', y='Min', color='Canal', barmode='group', text='Texto_hhmmss', color_discrete_sequence=['#2b7bba', '#7b3294'], height=290)
            fig3.update_layout(xaxis_title="", yaxis_title="minutos", showlegend=True, margin=dict(l=30, r=10, t=10, b=30), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"))
            fig3.update_traces(textposition='outside')
            st.markdown('<div class="chart-card">', unsafe_allow_html=True)
            st.plotly_chart(fig3, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("Valores de TME não detectados.")

# ==============================================================================
# 2. GUIA: NS POR OPERAÇÃO
# ==============================================================================
elif aba_selecionada == "NS por Operação":
    st.title("NS por Operação — Mensageria × Telefonia")
    
    df_ops = df_resumo[df_resumo['Operação'] != 'GERAL']
    df_melt = df_ops.melt(id_vars=['Operação'], value_vars=['NS Mensageria', 'NS Telefonia'], var_name='Canal', value_name='Val')
    df_melt['Porcentagem (%)'] = df_melt['Val'] * 100
    df_melt['Canal'] = df_melt['Canal'].map({'NS Mensageria': 'Mensageria', 'NS Telefonia': 'Telefonia'})
    
    fig_macro = px.bar(df_melt, x='Operação', y='Porcentagem (%)', color='Canal', barmode='group', text='Porcentagem (%)', color_discrete_sequence=['#2b7bba', '#7b3294'], height=350)
    fig_macro.update_yaxes(ticksuffix="%", range=[0, 100])
    fig_macro.update_layout(xaxis_title="", yaxis_title="", showlegend=True, legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"), margin=dict(t=10, b=10))
    fig_macro.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.plotly_chart(fig_macro, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    lista_ops = ['SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
    
    for i in range(0, len(lista_ops), 2):
        cols = st.columns(2)
        
        for j in range(2):
            if i + j < len(lista_ops):
                o = lista_ops[i + j]
                with cols[j]:
                    st.markdown(f'<div class="chart-card-compact">', unsafe_allow_html=True)
                    st.markdown(f"<p style='text-align: center; font-weight: bold; margin-bottom: 5px;'>{o}</p>", unsafe_allow_html=True)
                    
                    if o in df_diario['Operação'].values:
                        df_sub = df_diario[df_diario['Operação'] == o].sort_values('Dia')
                        fig_bar = go.Figure()
                        
                        fig_bar.add_trace(go.Bar(x=df_sub['Dia'], y=df_sub['NS Msg']*100, name='Mensageria', marker_color='#2b7bba'))
                        fig_bar.add_trace(go.Bar(x=df_sub['Dia'], y=df_sub['NS Voz']*100, name='Telefonia', marker_color='#7b3294'))
                        fig_bar.add_trace(go.Scatter(x=df_sub['Dia'], y=[90]*len(df_sub), name='Meta 90%', mode='lines', line=dict(color='gray', width=2, dash='dash')))
                        
                        fig_bar.update_layout(
                            barmode='group', 
                            height=280, 
                            margin=dict(l=25, r=10, t=10, b=20), 
                            showlegend=True,
                            legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
                            yaxis=dict(range=[0, 105], ticksuffix="%")
                        )
                        st.plotly_chart(fig_bar, use_container_width=True)
                    else:
                        st.info("Sem dados")
                    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 3. GUIA: NS DIÁRIO (PERDAS)
# ==============================================================================
elif aba_selecionada == "NS Diário (perdas)":
    st.title("NS Diário por Operação (perdas)")
    st.markdown("<p style='color: #657180; font-size: 14px; margin-top: -15px;'>NS por dia em cada operação. Vermelho = abaixo de 90% (perda).</p>", unsafe_allow_html=True)

    colunas_ordem = ['GERAL', 'SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
    colunas_display = ['Geral', 'SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
    df_chart_data = df_diario[df_diario['Operação'] != 'GERAL'].sort_values('Dia').copy()
    df_chart_data['NS Msg (%)'] = df_chart_data['NS Msg'] * 100
    df_chart_data['NS Voz (%)'] = df_chart_data['NS Voz'] * 100
    cores_ops = {'SAC': '#5b9bd5', 'RETENÇÃO': '#ffc000', 'COBRANÇA': '#c00000', 'SUPORTE': '#70ad47', 'MULTISKILL': '#7030a0'}

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Mensageria")
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        if not df_chart_data.empty:
            fig_loss_msg = px.line(df_chart_data, x='Dia', y='NS Msg (%)', color='Operação', color_discrete_map=cores_ops, height=350)
            dias_unicos = df_chart_data['Dia'].unique()
            fig_loss_msg.add_trace(go.Scatter(x=dias_unicos, y=[90]*len(dias_unicos), name='Meta 90%', line=dict(color='darkgray', width=1.5, dash='dash'), mode='lines'))
            fig_loss_msg.update_traces(mode='lines+markers', marker=dict(size=5), line_shape='spline')
            fig_loss_msg.update_layout(xaxis=dict(title="", tickangle=-90, type='category'), yaxis=dict(title="", ticksuffix="%", range=[0, 105], dtick=10), margin=dict(l=30, r=30, t=10, b=10), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center", title=""), plot_bgcolor='white')
            fig_loss_msg.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
            fig_loss_msg.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
            st.plotly_chart(fig_loss_msg, use_container_width=True)
            
            df_mat_msg = df_diario.pivot(index='Dia', columns='Operação', values='NS Msg').reset_index()
            cols_existentes = [col for col in colunas_ordem if col in df_mat_msg.columns]
            df_mat_msg = df_mat_msg[['Dia'] + cols_existentes]
            st.markdown(renderizar_tabela_html(df_mat_msg), unsafe_allow_html=True)
        else:
            st.warning("Sem dados diários processados.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.subheader("Telefonia")
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        if not df_chart_data.empty:
            fig_loss_voz = px.line(df_chart_data, x='Dia', y='NS Voz (%)', color='Operação', color_discrete_map=cores_ops, height=350)
            fig_loss_voz.add_trace(go.Scatter(x=dias_unicos, y=[90]*len(dias_unicos), name='Meta 90%', line=dict(color='darkgray', width=1.5, dash='dash'), mode='lines'))
            fig_loss_voz.update_traces(mode='lines+markers', marker=dict(size=5), line_shape='spline')
            fig_loss_voz.update_layout(xaxis=dict(title="", tickangle=-90, type='category'), yaxis=dict(title="", ticksuffix="%", range=[0, 105], dtick=10), margin=dict(l=30, r=30, t=10, b=10), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center", title=""), plot_bgcolor='white')
            fig_loss_voz.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
            fig_loss_voz.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
            st.plotly_chart(fig_loss_voz, use_container_width=True)
            
            df_mat_voz = df_diario.pivot(index='Dia', columns='Operação', values='NS Voz').reset_index()
            cols_existentes_voz = [col for col in colunas_ordem if col in df_mat_voz.columns]
            df_mat_voz = df_mat_voz[['Dia'] + cols_existentes_voz]
            st.markdown(renderizar_tabela_html(df_mat_voz), unsafe_allow_html=True)
        else:
            st.warning("Sem dados diários processados.")
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 4. GUIA: TABELA CONSOLIDADA
# ==============================================================================
elif aba_selecionada == "Tabela Consolidada":
    st.title("Tabela Resumo (NS, TMA e TME)")
    df_tab = df_resumo.copy()
    
    df_totais = df_diario[df_diario['Operação'] != 'GERAL'].groupby('Operação')[['Vol Msg', 'Vol Voz']].sum().reset_index()
    total_geral_msg = df_totais['Vol Msg'].sum()
    total_geral_voz = df_totais['Vol Voz'].sum()
    
    df_report = pd.DataFrame()
    df_report['Operação'] = df_tab['Operação']
    
    df_report['Total Msg'] = df_report['Operação'].map(df_totais.set_index('Operação')['Vol Msg'])
    df_report.loc[df_report['Operação'] == 'GERAL', 'Total Msg'] = total_geral_msg
    df_report['Total Msg'] = df_report['Total Msg'].fillna(0).map('{:,.0f}'.format).str.replace(',', '.')
    
    df_report['NS Msg'] = (df_tab['NS Mensageria'] * 100).map('{:.1f}%'.format)
    df_report['TMA Msg'] = df_tab['TMA Mensageria (Minutos)'].apply(formatar_para_ms)
    df_report['TME Msg'] = df_tab['TME Mensageria (Minutos)'].apply(formatar_para_ms) if 'TME Mensageria (Minutos)' in df_tab.columns else "—"
    
    df_report['Total Voz'] = df_report['Operação'].map(df_totais.set_index('Operação')['Vol Voz'])
    df_report.loc[df_report['Operação'] == 'GERAL', 'Total Voz'] = total_geral_voz
    df_report['Total Voz'] = df_report['Total Voz'].fillna(0).map('{:,.0f}'.format).str.replace(',', '.')
    
    df_report['NS Voz'] = (df_tab['NS Telefonia'] * 100).map('{:.1f}%'.format)
    df_report['TMA Voz'] = df_tab['TMA Telefonia (Minutos)'].apply(formatar_para_ms)
    
    tme_voz_col = 'TME Telefonia (Minutos)' if 'TME Telefonia (Minutos)' in df_tab.columns else ('TME (Minutos)' if 'TME (Minutos)' in df_tab.columns else None)
    df_report['TME Voz'] = df_tab[tme_voz_col].apply(formatar_para_ms) if tme_voz_col else "—"

    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.markdown(renderizar_tabela_ns_tma_html(df_report), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 5. GUIA: DIÁRIO POR OPERAÇÃO
# ==============================================================================
elif aba_selecionada == "Diário por Operação":
    st.title("Diário por Operação — Volume, NS e TMA")
    
    op_disponiveis = df_diario['Operação'].unique().tolist()
    if op_disponiveis:
        op_selecionada = st.selectbox("Selecione a Operação para Filtro:", op_disponiveis)
        df_filtro = df_diario[df_diario['Operação'] == op_selecionada].sort_values('Dia').copy()
        
        df_diario_rep = pd.DataFrame()
        df_diario_rep['Dia'] = df_filtro['Dia']
        df_diario_rep['Vol Msg'] = df_filtro['Vol Msg'].map('{:,.0f}'.format).str.replace(',', '.')
        df_diario_rep['NS Msg'] = (df_filtro['NS Msg'] * 100).map('{:.1f}%'.format)
        df_diario_rep['TMA Msg'] = df_filtro['TMA Msg (Min)'].apply(formatar_para_ms)
        df_diario_rep['Vol Voz'] = df_filtro['Vol Voz'].map('{:,.0f}'.format).str.replace(',', '.')
        df_diario_rep['NS Voz'] = (df_filtro['NS Voz'] * 100).map('{:.1f}%'.format)
        df_diario_rep['TMA Voz'] = df_filtro['TMA Voz (Min)'].apply(formatar_para_ms)

        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown(renderizar_tabela_diario_rep_html(df_diario_rep), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("Não há dados processados para exibir nesta aba.")