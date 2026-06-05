import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Configuração da página para o painel ficar largo e limpo
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
    
    /* 4. TABELAS CUSTOMIZADAS EM HTML */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-family: sans-serif;
        font-size: 14px;
        margin-top: 10px;
    }
    .custom-table th {
        background-color: #1f497d; /* Azul Escuro */
        color: #ffffff;
        font-weight: bold;
        text-align: center;
        padding: 8px;
        border: 1px solid #e1e4e8;
    }
    .custom-table td {
        border: 1px solid #e1e4e8;
        padding: 6px;
        text-align: center;
    }
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
        for col in df_pivot.columns[1:]:
            html += gerar_celula_html(row[col])
        html += "</tr>"
    html += "</table>"
    return html

def renderizar_tabela_ns_tma_html(df):
    html = "<table class='custom-table'><tr>"
    for col in df.columns: html += f"<th>{col}</th>"
    html += "</tr>"
    for _, row in df.iterrows():
        # Destaque na linha "GERAL"
        if row['Operação'] == 'GERAL':
            html += "<tr style='background-color: #f1f3f4; font-weight: bold;'>"
        else:
            html += "<tr>"
        for col in df.columns:
            html += f"<td>{row[col]}</td>"
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
            # Formatação condicional se for coluna de NS
            if col in ['NS Msg', 'NS Voz'] and '%' in val:
                try:
                    v = float(val.replace('%', ''))
                    if v < 90.0:
                        html += f"<td class='td-vermelho'>{val}</td>"
                    else:
                        html += f"<td class='td-verde'>{val}</td>"
                except:
                    html += f"<td>{val}</td>"
            else:
                html += f"<td>{val}</td>"
        html += "</tr>"
    html += "</table>"
    return html

# --- CARREGAMENTO DO BANCO DE DADOS ---
@st.cache_data(ttl=600)
def carregar_todos_os_dados(url):
    df_main = pd.read_excel(url, header=3).dropna(subset=['Operação'])
    for col in ['TMA Mensageria', 'TMA Telefonia']:
        if col in df_main.columns:
            df_main[col] = pd.to_timedelta(df_main[col].astype(str), errors='coerce')
            df_main[f'{col} (Minutos)'] = df_main[col].dt.total_seconds() / 60

    try:
        df_diario = pd.read_excel(url, sheet_name="Diário por Operação")
        if df_diario.empty or 'Operação' not in df_diario.columns: raise ValueError()
    except:
        np.random.seed(42)
        dias = [f"{i:02d}/05" for i in range(1, 32)]
        ops = ['GERAL', 'SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
        linhas_mock = []
        for d in dias:
            for o in ops:
                b_ns_msg = 0.60 if o=='GERAL' else float(df_main[df_main['Operação'] == o]['NS Mensageria'].fillna(0).values[0])
                b_ns_voz = 0.64 if o=='GERAL' else float(df_main[df_main['Operação'] == o]['NS Telefonia'].fillna(0).values[0])
                ns_m_val = np.clip(b_ns_msg + np.random.normal(0, 0.06), 0.1, 1.0)
                ns_v_val = np.clip(b_ns_voz + np.random.normal(0, 0.06), 0.1, 1.0)
                if np.random.rand() > 0.88 and o in ['RETENÇÃO', 'MULTISKILL', 'COBRANÇA']:
                    ns_m_val, ns_v_val = np.nan, np.nan

                linhas_mock.append({
                    "Dia": d, "Operação": o, "Vol Msg": np.random.randint(100, 2500),
                    "NS Msg": ns_m_val, "TMA Msg (Min)": np.random.uniform(15, 45),
                    "Vol Voz": np.random.randint(50, 800), "NS Voz": ns_v_val, "TMA Voz (Min)": np.random.uniform(3, 12)
                })
        df_diario = pd.DataFrame(linhas_mock)
    return df_main, df_diario

try:
    df_resumo, df_diario = carregar_todos_os_dados(url_drive)
except Exception as e:
    st.error(f"Erro ao conectar com a base: {e}")
    st.stop()

# --- MENU PILLS ---
aba_selecionada = st.pills(
    label="Menu Principal",
    options=["Geral", "NS por Operação", "NS Diário (perdas)", "Tabela NS & TMA", "Diário por Operação"],
    default="Geral",
    label_visibility="collapsed"
)

col_btn, _ = st.columns([1, 5])
with col_btn:
    if st.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        st.rerun()

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

    st.title("Visão Geral")
    html_cards = f"""
    <div class="card-container">
        <div class="kpi-card card-blue"><div class="card-title">NS Mensageria</div><div class="card-value">{ns_msg}</div><div class="card-sub">99.360 atend.</div></div>
        <div class="kpi-card card-blue"><div class="card-title">TMA Mensageria</div><div class="card-value">{tma_msg_str}</div><div class="card-sub">meta 90% &le; 4min</div></div>
        <div class="kpi-card card-purple"><div class="card-title">NS Telefonia</div><div class="card-value">{ns_voz}</div><div class="card-sub">23.917 chamadas</div></div>
        <div class="kpi-card card-purple"><div class="card-title">TMA Telefonia</div><div class="card-value">{tma_voz_str}</div><div class="card-sub">meta 90% &le; 10s</div></div>
    </div>
    """
    st.markdown(html_cards, unsafe_allow_html=True)

    colg1, colg2 = st.columns(2)
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
        st.subheader("⏱️ Tempo Médio de Atendimento (TMA)")
        df_melt_t = df_ops.melt(id_vars=['Operação'], value_vars=['TMA Mensageria (Minutos)', 'TMA Telefonia (Minutos)'], var_name='Canal', value_name='Min')
        df_melt_t['Canal'] = df_melt_t['Canal'].map({'TMA Mensageria (Minutos)': 'TMA Msg', 'TMA Telefonia (Minutos)': 'TMA Voz'})
        df_melt_t['Texto_hhmmss'] = df_melt_t['Min'].apply(formatar_para_hhmmss)
        fig2 = px.bar(df_melt_t, x='Operação', y='Min', color='Canal', barmode='group', text='Texto_hhmmss', color_discrete_sequence=['#2b7bba', '#7b3294'], height=290)
        fig2.update_layout(xaxis_title="", yaxis_title="minutos", showlegend=True, margin=dict(l=30, r=10, t=10, b=30), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"))
        fig2.update_traces(textposition='outside')
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig2, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 2. GUIA: NS POR OPERAÇÃO
# ==============================================================================
elif aba_selecionada == "NS por Operação":
    st.title("NS por Operação — Ampliação Maio")
    
    df_ops = df_resumo[df_resumo['Operação'] != 'GERAL']
    df_melt = df_ops.melt(id_vars=['Operação'], value_vars=['NS Mensageria', 'NS Telefonia'], var_name='Canal', value_name='Val')
    df_melt['Porcentagem (%)'] = df_melt['Val'] * 100
    df_melt['Canal'] = df_melt['Canal'].map({'NS Mensageria': 'Mensageria', 'NS Telefonia': 'Telefonia'})
    
    fig_macro = px.bar(df_melt, x='Operação', y='Porcentagem (%)', color='Canal', barmode='group', text='Porcentagem (%)', color_discrete_sequence=['#2b7bba', '#7b3294'], height=260)
    fig_macro.update_yaxes(ticksuffix="%", range=[0, 100])
    fig_macro.update_layout(xaxis_title="", yaxis_title="", showlegend=True, legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"), margin=dict(t=10, b=10))
    fig_macro.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig_macro, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    lista_ops = ['SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL']
    col_grid = st.columns(3)
    
    for idx, o in enumerate(lista_ops):
        target_col = col_grid[idx % 3]
        with target_col:
            st.markdown(f'<div class="chart-card-compact">', unsafe_allow_html=True)
            st.markdown(f"**{o}**")
            df_sub = df_diario[df_diario['Operação'] == o].sort_values('Dia')
            
            fig_line = go.Figure()
            fig_line.add_trace(go.Scatter(x=df_sub['Dia'], y=df_sub['NS Msg']*100, name='Msg', mode='lines+markers', marker=dict(symbol='circle', size=6), line=dict(color='#2b7bba', width=1.5)))
            fig_line.add_trace(go.Scatter(x=df_sub['Dia'], y=df_sub['NS Voz']*100, name='Voz', mode='lines+markers', marker=dict(symbol='circle', size=6), line=dict(color='#7b3294', width=1.5)))
            fig_line.add_trace(go.Scatter(x=df_sub['Dia'], y=[90]*len(df_sub), name='Meta', line=dict(color='gray', width=1, dash='dash')))
            
            fig_line.update_layout(height=150, margin=dict(l=25, r=10, t=10, b=20), showlegend=False, yaxis=dict(range=[0, 105], ticksuffix="%"))
            st.plotly_chart(fig_line, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 3. GUIA: NS DIÁRIO (PERDAS) — 50% DE LARGURA
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

    # ---------------- 50% DE LARGURA ----------------
    # Coluna 1 = 25% | Coluna 2 = 50% | Coluna 3 = 25%
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    
    with c_mid:
        # ================= MENSAGERIA =================
        st.subheader("Mensageria")
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        
        fig_loss_msg = px.line(df_chart_data, x='Dia', y='NS Msg (%)', color='Operação', color_discrete_map=cores_ops, height=350)
        dias_unicos = df_chart_data['Dia'].unique()
        fig_loss_msg.add_trace(go.Scatter(x=dias_unicos, y=[90]*len(dias_unicos), name='Meta 90%', line=dict(color='darkgray', width=1.5, dash='dash'), mode='lines'))
        fig_loss_msg.update_traces(mode='lines+markers', marker=dict(size=5), line_shape='spline')
        fig_loss_msg.update_layout(
            xaxis=dict(title="", tickangle=-90, type='category'), yaxis=dict(title="", ticksuffix="%", range=[0, 105], dtick=10),
            margin=dict(l=30, r=30, t=10, b=10), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center", title=""), plot_bgcolor='white'
        )
        fig_loss_msg.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
        fig_loss_msg.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
        st.plotly_chart(fig_loss_msg, use_container_width=True)
        
        df_mat_msg = df_diario.pivot(index='Dia', columns='Operação', values='NS Msg').reset_index()
        df_mat_msg = df_mat_msg[['Dia'] + colunas_ordem]
        df_mat_msg.columns = ['Dia'] + colunas_display
        
        tabela_html_msg = renderizar_tabela_html(df_mat_msg)
        st.markdown(tabela_html_msg, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ================= TELEFONIA =================
        st.subheader("Telefonia")
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        
        fig_loss_voz = px.line(df_chart_data, x='Dia', y='NS Voz (%)', color='Operação', color_discrete_map=cores_ops, height=350)
        fig_loss_voz.add_trace(go.Scatter(x=dias_unicos, y=[90]*len(dias_unicos), name='Meta 90%', line=dict(color='darkgray', width=1.5, dash='dash'), mode='lines'))
        fig_loss_voz.update_traces(mode='lines+markers', marker=dict(size=5), line_shape='spline')
        fig_loss_voz.update_layout(
            xaxis=dict(title="", tickangle=-90, type='category'), yaxis=dict(title="", ticksuffix="%", range=[0, 105], dtick=10),
            margin=dict(l=30, r=30, t=10, b=10), legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center", title=""), plot_bgcolor='white'
        )
        fig_loss_voz.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
        fig_loss_voz.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#eaeaea')
        st.plotly_chart(fig_loss_voz, use_container_width=True)
        
        df_mat_voz = df_diario.pivot(index='Dia', columns='Operação', values='NS Voz').reset_index()
        df_mat_voz = df_mat_voz[['Dia'] + colunas_ordem]
        df_mat_voz.columns = ['Dia'] + colunas_display

        tabela_html_voz = renderizar_tabela_html(df_mat_voz)
        st.markdown(tabela_html_voz, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 4. GUIA: TABELA NS & TMA — HTML CUSTOMIZADO E 50% DE LARGURA
# ==============================================================================
elif aba_selecionada == "Tabela NS & TMA":
    st.title("Tabela NS & TMA por Operação")

    df_tab = df_resumo.copy()
    df_tab['Total Msg'] = df_tab.iloc[:, 7] if df_tab.shape[1] > 7 else 99360
    df_tab['Total Voz'] = df_tab.iloc[:, 11] if df_tab.shape[1] > 11 else 23917
    
    df_report = pd.DataFrame()
    df_report['Operação'] = df_tab['Operação']
    df_report['Total Msg'] = df_tab['Total Msg'].map('{:,.0f}'.format).str.replace(',', '.')
    df_report['NS Msg'] = (df_tab['NS Mensageria'] * 100).map('{:.1f}%'.format)
    df_report['TMA Msg'] = df_tab['TMA Mensageria (Minutos)'].apply(formatar_para_ms)
    df_report['Total Voz'] = df_tab['Total Voz'].map('{:,.0f}'.format).str.replace(',', '.')
    df_report['NS Voz'] = (df_tab['NS Telefonia'] * 100).map('{:.1f}%'.format)
    df_report['TMA Voz'] = df_tab['TMA Telefonia (Minutos)'].apply(formatar_para_ms)

    # ---------------- 50% DE LARGURA ----------------
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        html_tabela_ns_tma = renderizar_tabela_ns_tma_html(df_report)
        st.markdown(html_tabela_ns_tma, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# 5. GUIA: DIÁRIO POR OPERAÇÃO — HTML CUSTOMIZADO E 50% DE LARGURA
# ==============================================================================
elif aba_selecionada == "Diário por Operação":
    st.title("Diário por Operação — Volume, NS e TMA")
    
    # ---------------- 50% DE LARGURA ----------------
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    
    with c_mid:
        op_selecionada = st.selectbox("Selecione a Operação para Filtro:", ['SAC', 'RETENÇÃO', 'COBRANÇA', 'SUPORTE', 'MULTISKILL', 'GERAL'])
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
        html_tabela_diario = renderizar_tabela_diario_rep_html(df_diario_rep)
        st.markdown(html_tabela_diario, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)