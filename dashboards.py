# --- START OF FILE dashboard_filtro_movimento_indice6.py ---
# Lembre-se de instalar a biblioteca para exportar para Excel, se ainda não tiver:
# pip install xlsxwriter

import streamlit as st
import pandas as pd
import plotly.express as px
import io # Para manipulação de bytes em memória (usado para PDF e Excel)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
import plotly.io as pio
# import numpy as np # Não estritamente necessário com as modificações atuais

MESES_PT_ORDENADOS = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
MESES_PT_MAP = {i+1: mes for i, mes in enumerate(MESES_PT_ORDENADOS)}

# Função para converter DataFrame para bytes de Excel
def df_to_excel_bytes(df_to_export):
    output = io.BytesIO()
    # Usar um with statement garante que o writer seja fechado corretamente
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Escreve o DataFrame na Planilha1, sem o índice do pandas
        df_to_export.to_excel(writer, index=False, sheet_name='Dados')
    processed_data = output.getvalue()
    return processed_data

st.set_page_config(layout="wide", page_title="Dashboard de Consumo com PDF")
# FUNÇÃO PARA CARREGAR DADOS DE ESTOQUE
@st.cache_data
def load_stock_data():
    try:
        df = pd.read_csv("estoqueatual.csv", sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv("estoqueatual.csv", sep=';', encoding='latin1')
    except FileNotFoundError:
        st.sidebar.warning("Arquivo 'estoqueatual.csv' não encontrado! A análise de estoque não será exibida.")
        return pd.DataFrame()
    except Exception as e:
        st.sidebar.error(f"Erro ao ler CSV de estoque: {e}")
        return pd.DataFrame()

    # Validação e preparação das colunas
    if "Insumo" not in df.columns or "Quantidade" not in df.columns:
        st.sidebar.error("O arquivo 'estoqueatual.csv' deve conter as colunas 'Insumo' e 'Quantidade'.")
        return pd.DataFrame()
    
    # Garante que a chave de junção é do tipo string
    df['Insumo'] = df['Insumo'].astype(str)
    
    # Converte a coluna "Quantidade" para numérico
    df['Quantidade'] = pd.to_numeric(df['Quantidade'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)
    df['Quantidade'] = df['Quantidade'].apply(lambda x: x if x > 0 else 0)

    # Renomeia "Quantidade" para "Estoque Atual" para o resto do script funcionar
    df.rename(columns={'Quantidade': 'Estoque Atual'}, inplace=True)

    # Retorna apenas as colunas necessárias
    return df[['Insumo', 'Estoque Atual']]


@st.cache_data
def load_data():
    try:
        df = pd.read_csv("Material-CSVANUAL.csv", sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv("Material-CSVANUAL.csv", sep=';', encoding='latin1')
    except FileNotFoundError: st.error("Arquivo 'Material-CSVANUAL.csv' não encontrado."); return pd.DataFrame()
    except pd.errors.EmptyDataError: st.error("Arquivo 'Material-CSVANUAL.csv' está vazio."); return pd.DataFrame()
    except Exception as e: st.error(f"Erro ao ler CSV: {e}"); return pd.DataFrame()
    return df

@st.cache_data
def preprocess_data(df_original):
    if df_original.empty: return pd.DataFrame()
    df = df_original.copy()
    column_mapping = {
        'Insumo': 'Cód. Insumo', 'Descricao': 'Desc. Insumo', 'Dt Movimento': 'Dt Movimento',
        'Quantidade': 'Quantidade', 'Descricao Movimento': 'Descricao Movimento',
        'Descricao Requisitante': 'Descricao Requisitante', 'Valor ': 'Valor',
        'Descricao da Classe': 'Descricao Classe',
        'Fornecedor': 'Nome Fornecedor'
    }
    actual_renames = {}
    for original_name_csv, new_name_internal in column_mapping.items():
        if original_name_csv in df.columns: actual_renames[original_name_csv] = new_name_internal
        else: st.sidebar.warning(f"Coluna original '{original_name_csv}' do CSV não encontrada para mapeamento. Será ignorada.")
    df.rename(columns=actual_renames, inplace=True)

    if 'Quantidade' not in df.columns: st.error("Coluna interna 'Quantidade' não encontrada."); return pd.DataFrame()
    df['Quantidade'] = pd.to_numeric(df['Quantidade'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').abs()
    if 'Dt Movimento' not in df.columns: st.error("Coluna interna 'Dt Movimento' não encontrada."); return pd.DataFrame()
    df['Dt Movimento'] = pd.to_datetime(df['Dt Movimento'], dayfirst=True, errors='coerce')
    df['Ano'] = df['Dt Movimento'].dt.year
    df['Mês Num'] = df['Dt Movimento'].dt.month
    df['Mês Nome'] = df['Mês Num'].map(MESES_PT_MAP)
    
    df = df.dropna(subset=['Dt Movimento', 'Ano', 'Mês Num', 'Mês Nome']); 
    df['Ano'] = df['Ano'].astype(int)
    df['Mês Num'] = df['Mês Num'].astype(int)

    essential_cols_internal = ['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento', 'Quantidade', 'Ano', 'Dt Movimento', 'Mês Num', 'Mês Nome']
    optional_cols_internal = ['Descricao Requisitante', 'Valor', 'Descricao Classe', 'Nome Fornecedor'] 
    
    for col in essential_cols_internal:
        if col not in df.columns: st.error(f"Coluna interna essencial '{col}' não encontrada."); return pd.DataFrame()
    for col in optional_cols_internal:
        if col not in df.columns: 
            if col == 'Valor': df[col] = 0.0
            elif col == 'Nome Fornecedor': df[col] = 'N/A' 
            else: df[col] = 'N/A' 
    
    df = df.dropna(subset=['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento'])
    df['Desc. Insumo'] = df['Desc. Insumo'].astype(str); df['Cód. Insumo'] = df['Cód. Insumo'].astype(str)
    if 'Descricao Requisitante' in df.columns: df['Descricao Requisitante'] = df['Descricao Requisitante'].astype(str).fillna('N/A')
    if 'Descricao Classe' in df.columns: df['Descricao Classe'] = df['Descricao Classe'].astype(str).fillna('N/A')
    if 'Valor' in df.columns: df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').abs().fillna(0)
    if 'Nome Fornecedor' in df.columns: df['Nome Fornecedor'] = df['Nome Fornecedor'].astype(str).fillna('N/A')

    if 'Nome Fornecedor' in df.columns:
        valores_fornecedor_a_remover = ["AJUSTE DE INVENTARIO", "AJUSTE INVENTARIO", "AJUSTE MATERIAL"]
        df = df[~df['Nome Fornecedor'].isin(valores_fornecedor_a_remover)]
    return df

def generate_pdf_report(
    selected_desc_insumos_pdf, selected_cod_insumos_pdf, selected_years_pdf,
    selected_movimento_consumo_pdf, selected_classes_pdf,
    consumo_anual_pivot_df, consumo_mensal_pivot_df, fig_consumo_anual_line_obj, fig_consumo_mensal_bar_obj,
    media_geral_mensal_df,
    consumo_mensal_detalhado_pdf_data, fig_consumo_mensal_detalhado_obj, 
    material_analise_unidade_pdf=None, material_analise_unidade_cod_insumo_pdf=None,
    media_mensal_unidade_df=None, fig_unidade_media_obj=None, pivot_unidade_media_mensal_df=None
    ):
    buffer = io.BytesIO(); doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), margins=[inch/2]*4); styles = getSampleStyleSheet(); story = []
    story.append(Paragraph("Relatório de Análise de Consumo de Materiais", styles['h1'])); story.append(Spacer(1, 0.2*inch))
    if selected_desc_insumos_pdf: story.append(Paragraph(f"<b>Descrições de Insumos:</b> {', '.join(selected_desc_insumos_pdf)}", styles['Normal']))
    if selected_cod_insumos_pdf: story.append(Paragraph(f"<b>Códigos de Insumos:</b> {', '.join(selected_cod_insumos_pdf)}", styles['Normal']))
    if selected_classes_pdf: story.append(Paragraph(f"<b>Classes Selecionadas:</b> {', '.join(selected_classes_pdf)}", styles['Normal']))
    if not selected_desc_insumos_pdf and not selected_cod_insumos_pdf and not selected_classes_pdf: story.append(Paragraph("<b>Filtros de Insumo/Classe:</b> Nenhum aplicado", styles['Normal']))
    story.append(Paragraph(f"<b>Anos:</b> {', '.join(map(str, selected_years_pdf)) if selected_years_pdf else 'Nenhum'}", styles['Normal']))
    story.append(Paragraph(f"<b>Tipo de Movimento:</b> {selected_movimento_consumo_pdf}", styles['Normal'])); story.append(Spacer(1, 0.2*inch))

    def df_to_table(df, title=""):
        elements_for_keeptogether = []
        if title: elements_for_keeptogether.extend([Paragraph(f"<b>{title}</b>", styles['h3']), Spacer(1, 0.1*inch)])
        if df.empty: elements_for_keeptogether.extend([Paragraph("Nenhum dado para exibir.", styles['Italic']), Spacer(1, 0.1*inch)]); story.append(KeepTogether(elements_for_keeptogether)); return
        
        max_cols = 15 
        df_display = df.copy()
        if isinstance(df_display.columns, pd.MultiIndex):
            df_display.columns = ['_'.join(map(str, col)).strip('_') for col in df_display.columns.values]

        for col_name in df_display.columns:
            if pd.api.types.is_numeric_dtype(df_display[col_name]):
                if col_name == 'CONSUMO MEDIO': # Nome da coluna como string
                    df_display[col_name] = df_display[col_name].apply(lambda x: f"{x:,.1f}" if pd.notnull(x) else '0.0')
                elif col_name not in ['CODIGO', 'DESCRICAO']: 
                    df_display[col_name] = df_display[col_name].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else '0')

        if len(df_display.columns) > max_cols: 
            df_display_subset = df_display.iloc[:, :max_cols].copy()
            elements_for_keeptogether.append(Paragraph(f"(Exibindo as primeiras {max_cols} colunas de {len(df_display.columns)}.)", styles['Italic']))
            data = [df_display_subset.columns.to_list()] + df_display_subset.astype(str).values.tolist()
        else:
            data = [df_display.columns.to_list()] + df_display.astype(str).values.tolist()

        table = Table(data, repeatRows=1); table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),('FONTSIZE', (0,0), (-1,-1), 7), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elements_for_keeptogether.append(table)
        elements_for_keeptogether.append(Spacer(1, 0.2*inch))
        story.append(KeepTogether(elements_for_keeptogether))

    def fig_to_image_reportlab(fig, title=""):
        elements_for_keeptogether = []
        if title: elements_for_keeptogether.extend([Paragraph(f"<b>{title}</b>", styles['h3']), Spacer(1, 0.1*inch)])
        if fig is None: elements_for_keeptogether.extend([Paragraph("Gráfico não disponível.", styles['Italic']), Spacer(1, 0.1*inch)]); story.append(KeepTogether(elements_for_keeptogether)); return
        try:
            img_bytes = pio.to_image(fig, format="png", width=700, height=350, scale=1.5) 
            img = Image(io.BytesIO(img_bytes), width=6.8*inch, height=3.4*inch) 
            img.hAlign = 'CENTER'
            elements_for_keeptogether.append(img)
        except Exception as e: 
            error_msg = f"Erro renderizar gráfico '{title}' PDF: {str(e)}"
            st.sidebar.error(error_msg)
            elements_for_keeptogether.append(Paragraph(error_msg, styles['Italic']))
        
        elements_for_keeptogether.append(Spacer(1, 0.1*inch)) 
        story.append(KeepTogether(elements_for_keeptogether))
    
    story.append(Paragraph("Análise de Consumo por Insumo", styles['h2']))
    df_to_table(consumo_anual_pivot_df, "Consumo Total Anual por Insumo")
    df_to_table(consumo_mensal_pivot_df, "Consumo Médio Mensal Agregado por Ano (calculado sobre meses com consumo)") 
    fig_to_image_reportlab(fig_consumo_anual_line_obj, "Tendência de Consumo Total Anual")
    fig_to_image_reportlab(fig_consumo_mensal_bar_obj, "Comparativo de Consumo Médio Mensal Agregado por Ano (calculado sobre meses com consumo)")

    story.append(PageBreak()) 
    story.append(Paragraph("Análise Detalhada de Consumo Mensal", styles['h2']))
    pdf_table_title = "Consumo Mensal Efetivo por Insumo"
    if consumo_mensal_detalhado_pdf_data is not None and not consumo_mensal_detalhado_pdf_data.empty:
        if len(selected_years_pdf) == 1:
            pdf_table_title = f"Consumo Mensal Efetivo por Insumo (Ano: {selected_years_pdf[0]})"
        elif len(selected_years_pdf) > 1:
            pdf_table_title = f"Consumo Mensal Efetivo por Insumo (Média entre Anos: {', '.join(map(str,selected_years_pdf))})"
        df_to_table(consumo_mensal_detalhado_pdf_data, pdf_table_title)
    else:
        story.append(Paragraph("Dados de consumo mensal detalhado não disponíveis.", styles['Normal']))
    fig_to_image_reportlab(fig_consumo_mensal_detalhado_obj, "Gráfico de Consumo Mensal Efetivo por Insumo")

    story.append(PageBreak()) 
    story.append(Paragraph("Média Geral Mensal de Consumo por Insumo (sobre os anos selecionados)", styles['h2']))
    df_to_table(media_geral_mensal_df, "Média Geral Mensal por Insumo (baseada na média do consumo anual / nº meses com consumo)") 

    if material_analise_unidade_pdf and media_mensal_unidade_df is not None and not media_mensal_unidade_df.empty:
        story.append(PageBreak())
        titulo_unidade = f"Análise de Consumo por Unidade para o Insumo: {material_analise_unidade_pdf}"
        if material_analise_unidade_cod_insumo_pdf:
            titulo_unidade += f" (Cód: {material_analise_unidade_cod_insumo_pdf})"
        story.append(Paragraph(titulo_unidade, styles['h2']))
        
        df_to_table(media_mensal_unidade_df, f"Média Mensal de Consumo por Unidade (calculado sobre meses com consumo)")
        if fig_unidade_media_obj: fig_to_image_reportlab(fig_unidade_media_obj, f"Top Unidades por Média Mensal de Consumo (calculado sobre meses com consumo)")
        if pivot_unidade_media_mensal_df is not None and not pivot_unidade_media_mensal_df.empty: df_to_table(pivot_unidade_media_mensal_df, "Detalhe: Média Mensal por Unidade/Ano (calculado sobre meses com consumo)")
    doc.build(story); buffer.seek(0); return buffer.getvalue()

# --- Carregar e pré-processar os dados ---
raw_df = load_data()
material_df = preprocess_data(raw_df) 
df_estoque = load_stock_data() # ADICIONE ESTA LINHA

# --- Interface do Dashboard ---
st.title("📊 Dashboard Avançado de Análise de Consumo")

fig_consumo_anual_line = None; fig_consumo_mensal_bar = None; fig_media_geral_mensal_grafico = None; fig_unidade_media = None
fig_consumo_mensal_detalhado = None 
consumo_anual_pivot_pdf = pd.DataFrame(); consumo_mensal_pivot_pdf = pd.DataFrame()
media_geral_mensal_pdf = pd.DataFrame()
consumo_mensal_detalhado_pdf_display = pd.DataFrame() 
material_para_analise_unidade_global = None
codigo_insumo_para_unidade_global = "" 
media_mensal_por_unidade_pdf = pd.DataFrame(); pivot_unidade_ano_media_mensal_pdf = pd.DataFrame()

if material_df.empty: st.warning("Dados não carregados/processados adequadamente. Verifique o CSV e o mapeamento de colunas."); st.stop()

st.sidebar.header("⚙️ Filtros de Análise")
all_desc_insumos = sorted(material_df['Desc. Insumo'].dropna().unique())
selected_desc_insumos = st.sidebar.multiselect("💊 Selecione Insumos por Descrição:", options=all_desc_insumos, default=[])
all_cod_insumos = sorted(material_df['Cód. Insumo'].dropna().unique())
selected_cod_insumos = st.sidebar.multiselect("🔢 Selecione Insumos por Código:", options=all_cod_insumos, default=[])
all_classes = [];
if 'Descricao Classe' in material_df.columns: all_classes = sorted(material_df['Descricao Classe'].dropna().unique())
selected_classes = st.sidebar.multiselect("🏷️ Selecione Classes:", options=all_classes, default=[])
all_years = sorted(material_df['Ano'].dropna().unique())
selected_years = st.sidebar.multiselect("📅 Selecione os Anos:", options=all_years, default=[])

movimento_options = sorted(material_df['Descricao Movimento'].dropna().unique()) 
default_movimento_index = 0 
if len(movimento_options) > 6: default_movimento_index = 6
elif movimento_options: default_movimento_index = 0
selected_movimento_consumo = st.sidebar.selectbox("📉 Tipo de Movimento para Consumo:", options=movimento_options, index=default_movimento_index if movimento_options else 0)
pdf_download_button_placeholder = st.sidebar.empty()

condition_desc = material_df['Desc. Insumo'].isin(selected_desc_insumos) if selected_desc_insumos else pd.Series(True, index=material_df.index)
condition_cod = material_df['Cód. Insumo'].isin(selected_cod_insumos) if selected_cod_insumos else pd.Series(True, index=material_df.index)
if selected_desc_insumos and selected_cod_insumos: combined_insumo_condition = condition_desc | condition_cod
elif selected_desc_insumos: combined_insumo_condition = condition_desc
elif selected_cod_insumos: combined_insumo_condition = condition_cod
else: combined_insumo_condition = pd.Series(True, index=material_df.index)
condition_classe = pd.Series(True, index=material_df.index)
if selected_classes and 'Descricao Classe' in material_df.columns: condition_classe = material_df['Descricao Classe'].isin(selected_classes)
df_insumos_selecionados_base = material_df[combined_insumo_condition & condition_classe]
actual_selected_insumo_descriptions = sorted(df_insumos_selecionados_base['Desc. Insumo'].unique()) if not df_insumos_selecionados_base.empty else []

proceed_with_analysis = True
if not selected_years: st.info("👈 Por favor, selecione pelo menos um ano."); proceed_with_analysis = False
if not selected_movimento_consumo and movimento_options: st.info("👈 Por favor, selecione um tipo de movimento."); proceed_with_analysis = False
elif not selected_movimento_consumo and not movimento_options: st.error("Nenhum tipo de movimento disponível nos dados (após filtros)."); proceed_with_analysis = False 
if not actual_selected_insumo_descriptions and (selected_desc_insumos or selected_cod_insumos or selected_classes):
    st.warning("Nenhum insumo encontrado para a combinação de filtros de descrição, código e/ou classe selecionados.")
    proceed_with_analysis = False

if proceed_with_analysis:
    analysis_df_materiais = df_insumos_selecionados_base[
        (df_insumos_selecionados_base['Ano'].isin(selected_years)) &
        (df_insumos_selecionados_base['Descricao Movimento'] == selected_movimento_consumo)]
    if analysis_df_materiais.empty: st.warning(f"Nenhum dado encontrado para os critérios finais de filtro.")
    else:
        st.header("🔬 Análise de Consumo por Insumo")
        
        consumo_total_anual_df = analysis_df_materiais.groupby(
            ['Cód. Insumo', 'Desc. Insumo', 'Ano']
        )['Quantidade'].sum().reset_index()
        consumo_total_anual_df.rename(columns={'Quantidade': 'Consumo Total Anual'}, inplace=True)

        meses_com_consumo_df = analysis_df_materiais[analysis_df_materiais['Quantidade'] > 0].groupby(
            ['Cód. Insumo', 'Desc. Insumo', 'Ano']
        )['Mês Num'].nunique().reset_index()
        meses_com_consumo_df.rename(columns={'Mês Num': 'Nº Meses com Consumo'}, inplace=True)

        consumo_anual_por_material = pd.merge(
            consumo_total_anual_df,
            meses_com_consumo_df,
            on=['Cód. Insumo', 'Desc. Insumo', 'Ano'],
            how='left'
        )
        consumo_anual_por_material['Nº Meses com Consumo'] = consumo_anual_por_material['Nº Meses com Consumo'].fillna(0)
        
        consumo_anual_por_material['Consumo Médio Mensal (agregado)'] = consumo_anual_por_material.apply(
            lambda row: row['Consumo Total Anual'] / row['Nº Meses com Consumo'] if row['Nº Meses com Consumo'] > 0 else 0,
            axis=1
        )
        consumo_anual_por_material = consumo_anual_por_material.sort_values(by=['Cód. Insumo', 'Desc. Insumo', 'Ano'])
        
        st.subheader("Consumo Total Anual")
        try:
            consumo_anual_pivot_pdf = consumo_anual_por_material.pivot_table(index=['Cód. Insumo', 'Desc. Insumo'], columns='Ano', values='Consumo Total Anual', fill_value=0).reset_index()
            st.dataframe(consumo_anual_pivot_pdf.style.format({year: "{:,.0f}" for year in selected_years}), use_container_width=True)
            if not consumo_anual_pivot_pdf.empty:
                excel_data_cta = df_to_excel_bytes(consumo_anual_pivot_pdf)
                st.download_button(
                    label="📥 Exportar Consumo Total Anual para Excel",
                    data=excel_data_cta,
                    file_name="consumo_total_anual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e: st.error(f"Erro ao criar tabela de consumo anual: {str(e)}"); consumo_anual_pivot_pdf = pd.DataFrame()

        st.subheader("Consumo Médio Mensal (agregado por ano, calculado sobre meses com consumo)")
        try:
            consumo_mensal_pivot_pdf = consumo_anual_por_material.pivot_table(index=['Cód. Insumo', 'Desc. Insumo'], columns='Ano', values='Consumo Médio Mensal (agregado)', fill_value=0).reset_index()
            st.dataframe(consumo_mensal_pivot_pdf.style.format({year: "{:,.1f}" for year in selected_years}), use_container_width=True)
            if not consumo_mensal_pivot_pdf.empty:
                excel_data_cma = df_to_excel_bytes(consumo_mensal_pivot_pdf)
                st.download_button(
                    label="📥 Exportar Consumo Médio Mensal (agregado) para Excel",
                    data=excel_data_cma,
                    file_name="consumo_medio_mensal_agregado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e: st.error(f"Erro ao criar tabela de consumo mensal agregada: {str(e)}"); consumo_mensal_pivot_pdf = pd.DataFrame()

        if not consumo_anual_por_material.empty:
            fig_consumo_anual_line = px.line(consumo_anual_por_material, x='Ano', y='Consumo Total Anual', color='Desc. Insumo', markers=True, title='Tendência de Consumo Total Anual por Insumo', labels={'Desc. Insumo': 'Insumo'}, hover_data=['Cód. Insumo']); st.plotly_chart(fig_consumo_anual_line.update_layout(xaxis_type='category'), use_container_width=True)
            fig_consumo_mensal_bar = px.bar(consumo_anual_por_material, x='Ano', y='Consumo Médio Mensal (agregado)', color='Desc. Insumo', barmode='group', title='Comparativo de Consumo Médio Mensal (agregado por ano, calculado sobre meses com consumo)', labels={'Desc. Insumo': 'Insumo'}, hover_data=['Cód. Insumo']); st.plotly_chart(fig_consumo_mensal_bar.update_layout(xaxis_type='category'), use_container_width=True)

        st.markdown("---")
        st.header("📈 Análise Detalhada de Consumo Mensal")
        
        consumo_mensal_detalhado_calculado = analysis_df_materiais.groupby(
            ['Cód. Insumo', 'Desc. Insumo', 'Ano', 'Mês Num', 'Mês Nome']
        )['Quantidade'].sum().reset_index()

        if len(selected_years) > 1:
            consumo_mensal_grafico_df = consumo_mensal_detalhado_calculado.groupby(
                ['Cód. Insumo', 'Desc. Insumo', 'Mês Num', 'Mês Nome']
            )['Quantidade'].mean().reset_index()
            titulo_detalhado = f"Consumo Mensal Efetivo por Insumo (Média entre Anos: {', '.join(map(str,selected_years))})"
        else:
            consumo_mensal_grafico_df = consumo_mensal_detalhado_calculado.copy()
            titulo_detalhado = f"Consumo Mensal Efetivo por Insumo (Ano: {selected_years[0] if selected_years else 'N/A'})"
        
        consumo_mensal_grafico_df = consumo_mensal_grafico_df.sort_values(by=['Cód. Insumo', 'Desc. Insumo', 'Mês Num'])
        
        df_para_exibir_pivotado = pd.DataFrame()
        final_month_col_names = [] 

        if not consumo_mensal_grafico_df.empty:
            pivot_df = consumo_mensal_grafico_df.pivot_table(
                index=['Cód. Insumo', 'Desc. Insumo'],
                columns='Mês Nome',
                values='Quantidade',
                fill_value=0
            ).reset_index()

            pivot_df.rename(columns={'Cód. Insumo': 'CODIGO', 'Desc. Insumo': 'DESCRICAO'}, inplace=True)
            ordered_month_cols_orig = [mes for mes in MESES_PT_ORDENADOS if mes in pivot_df.columns]
            mapa_nomes_meses_novo = {}
            current_month_column_names = list(ordered_month_cols_orig)

            if len(selected_years) == 1:
                ano_selecionado_str = str(selected_years[0])
                ano_curto = ano_selecionado_str[-2:]
                temp_final_month_col_names = []
                for mes_longo in ordered_month_cols_orig:
                    mes_curto = mes_longo[:3].lower()
                    novo_nome = f"{mes_curto}/{ano_curto}"
                    mapa_nomes_meses_novo[mes_longo] = novo_nome
                    temp_final_month_col_names.append(novo_nome)
                pivot_df.rename(columns=mapa_nomes_meses_novo, inplace=True)
                current_month_column_names = temp_final_month_col_names
            
            final_month_col_names = current_month_column_names
            static_cols = ['CODIGO', 'DESCRICAO']
            cols_para_reordenar = static_cols + [col for col in final_month_col_names if col in pivot_df.columns]
            pivot_df = pivot_df[cols_para_reordenar]

            if len(selected_years) == 1:
                ano_selecionado = selected_years[0]
                media_ref = consumo_anual_por_material[
                    consumo_anual_por_material['Ano'] == ano_selecionado
                ][['Cód. Insumo', 'Desc. Insumo', 'Consumo Médio Mensal (agregado)']].copy()
                media_ref.rename(columns={
                    'Cód. Insumo': 'CODIGO', 
                    'Desc. Insumo': 'DESCRICAO',
                    'Consumo Médio Mensal (agregado)': 'CONSUMO MEDIO'
                    }, inplace=True)
                df_para_exibir_pivotado = pd.merge(pivot_df, media_ref, on=['CODIGO', 'DESCRICAO'], how='left')
                if 'CONSUMO MEDIO' in df_para_exibir_pivotado.columns:
                    df_para_exibir_pivotado['CONSUMO MEDIO'] = df_para_exibir_pivotado['CONSUMO MEDIO'].fillna(0.0)
                else: 
                    df_para_exibir_pivotado['CONSUMO MEDIO'] = 0.0
            else: 
                if final_month_col_names: 
                    # Garante que estamos somando apenas colunas que realmente existem e são numéricas (meses)
                    colunas_meses_existentes_no_pivot = [col for col in final_month_col_names if col in pivot_df.columns]
                    if colunas_meses_existentes_no_pivot:
                        pivot_df['Total Efetivo Agregado'] = pivot_df[colunas_meses_existentes_no_pivot].sum(axis=1)
                        pivot_df['Nº Meses com Consumo Efetivo Agregado'] = (pivot_df[colunas_meses_existentes_no_pivot] > 0).sum(axis=1)
                        pivot_df['CONSUMO MEDIO'] = pivot_df.apply(
                            lambda row: row['Total Efetivo Agregado'] / row['Nº Meses com Consumo Efetivo Agregado'] 
                                        if row['Nº Meses com Consumo Efetivo Agregado'] > 0 else 0,
                            axis=1
                        )
                    else: # Caso não haja colunas de meses válidas para somar
                         pivot_df['CONSUMO MEDIO'] = 0.0
                    cols_to_drop = [col for col in ['Total Efetivo Agregado', 'Nº Meses com Consumo Efetivo Agregado'] if col in pivot_df.columns]
                    df_para_exibir_pivotado = pivot_df.drop(columns=cols_to_drop) if cols_to_drop else pivot_df.copy()
                else: 
                    pivot_df['CONSUMO MEDIO'] = 0.0
                    df_para_exibir_pivotado = pivot_df.copy()
            
            if 'CONSUMO MEDIO' not in df_para_exibir_pivotado.columns:
                df_para_exibir_pivotado['CONSUMO MEDIO'] = 0.0
        
        else: 
            base_cols = ['CODIGO', 'DESCRICAO']
            month_cols_for_empty_df = []
            if len(selected_years) == 1:
                ano_s = str(selected_years[0])[-2:]
                month_cols_for_empty_df = [f"{m[:3].lower()}/{ano_s}" for m in MESES_PT_ORDENADOS]
            else:
                month_cols_for_empty_df = list(MESES_PT_ORDENADOS)
            
            final_month_col_names = month_cols_for_empty_df
            df_para_exibir_pivotado = pd.DataFrame(columns=base_cols + month_cols_for_empty_df + ['CONSUMO MEDIO'])
            for col in df_para_exibir_pivotado.columns:
                if col not in base_cols:
                    df_para_exibir_pivotado[col] = pd.Series(dtype='float64')
        
        consumo_mensal_detalhado_pdf_display = df_para_exibir_pivotado
        
        format_dict = {'CONSUMO MEDIO': "{:,.1f}"}
        for mes_col in final_month_col_names:
            if mes_col in df_para_exibir_pivotado.columns:
                format_dict[mes_col] = "{:,.0f}"
        
        st.subheader("Tabela: " + titulo_detalhado)
        if not df_para_exibir_pivotado.empty:
            st.dataframe(df_para_exibir_pivotado.style.format(format_dict, na_rep='0'), use_container_width=True)
            excel_data_cmd = df_to_excel_bytes(df_para_exibir_pivotado)
            st.download_button(
                label="📥 Exportar Consumo Mensal Detalhado para Excel",
                data=excel_data_cmd,
                file_name="consumo_mensal_detalhado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Nenhum dado detalhado de consumo mensal para exibir no formato pivotado.")

        if not consumo_mensal_grafico_df.empty:
            fig_consumo_mensal_detalhado = px.line( 
                consumo_mensal_grafico_df,
                x='Mês Nome', 
                y='Quantidade',
                color='Desc. Insumo',
                markers=True,
                title=titulo_detalhado,
                labels={'Desc. Insumo': 'Insumo', 'Quantidade': 'Consumo Mensal', 'Mês Nome': 'Mês'},
                hover_data=['Cód. Insumo'],
                category_orders={"Mês Nome": MESES_PT_ORDENADOS} 
            )
            st.plotly_chart(fig_consumo_mensal_detalhado, use_container_width=True)
        else:
            fig_consumo_mensal_detalhado = None 
        
        st.markdown("---")
       # COLE O NOVO BLOCO DE CÓDIGO AQUI

    st.markdown("---")
    
    # A partir daqui começa a nova lógica
    media_geral_mensal = consumo_anual_por_material.groupby(['Cód. Insumo', 'Desc. Insumo'])['Consumo Médio Mensal'].mean().reset_index()
    col_media_consumo_nome = f'Média Geral Mensal ({len(selected_years)}a)'
    media_geral_mensal.rename(columns={'Consumo Médio Mensal': col_media_consumo_nome}, inplace=True)
    
    st.header("📊 Análise de Cobertura e Status do Estoque")

    if not df_estoque.empty:
        st.sidebar.header("⚙️ Parâmetros de Estoque")
        critical_threshold = st.sidebar.number_input("Estoque Crítico (meses) <=", min_value=0.5, max_value=5.0, value=1.5, step=0.5)
        warning_threshold = st.sidebar.number_input("Alerta de Estoque (meses) <=", min_value=1.0, max_value=10.0, value=3.0, step=0.5)
        target_coverage = st.sidebar.number_input("Cobertura Ideal (meses)", min_value=2.0, max_value=24.0, value=6.0, step=1.0)
        
        st.caption(f"Status: Crítico <= {critical_threshold} meses | Alerta <= {warning_threshold} meses | Saudável > {warning_threshold} meses")

        df_estoque.rename(columns={'Insumo': 'Cód. Insumo'}, inplace=True)
        df_analise_estoque = pd.merge(media_geral_mensal, df_estoque, on='Cód. Insumo', how='left')
        df_analise_estoque['Estoque Atual'].fillna(0, inplace=True)

        def calcular_cobertura(row):
            media_consumo = row[col_media_consumo_nome]
            if media_consumo > 0:
                return row['Estoque Atual'] / media_consumo
            return float('inf')

        df_analise_estoque['Cobertura (Meses)'] = df_analise_estoque.apply(calcular_cobertura, axis=1)

        def definir_status(cobertura):
            if cobertura <= critical_threshold:
                return "Crítico"
            elif cobertura <= warning_threshold:
                return "Alerta"
            else:
                return "Saudável"
            
        df_analise_estoque['Status'] = df_analise_estoque['Cobertura (Meses)'].apply(definir_status)
        
        def calcular_pedido(row):
            estoque_ideal = row[col_media_consumo_nome] * target_coverage
            qtd_a_pedir = estoque_ideal - row['Estoque Atual']
            
            if qtd_a_pedir > 0 and row['Status'] != 'Saudável':
                return round(qtd_a_pedir, 0)
            return 0

        df_analise_estoque['Qtd. Sugerida p/ Pedido'] = df_analise_estoque.apply(calcular_pedido, axis=1)

        df_display_estoque = df_analise_estoque[[
            'Cód. Insumo', 'Desc. Insumo', col_media_consumo_nome, 
            'Estoque Atual', 'Cobertura (Meses)', 'Status', 'Qtd. Sugerida p/ Pedido'
        ]]

        format_dict_estoque = {
            col_media_consumo_nome: "{:,.1f}",
            'Estoque Atual': "{:,.0f}",
            'Cobertura (Meses)': lambda x: f"{x:,.1f}" if x != float('inf') else "N/A (sem consumo)",
            'Qtd. Sugerida p/ Pedido': "{:,.0f}"
        }

        st.dataframe(df_display_estoque.style.format(format_dict_estoque), use_container_width=True)
        
        excel_data_estoque = df_to_excel_bytes(df_display_estoque)
        st.download_button(
            label="📥 Exportar Análise de Estoque para Excel",
            data=excel_data_estoque,
            file_name="analise_cobertura_estoque.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.subheader("⚖️ Média Geral Mensal de Consumo")
        st.caption("Arquivo 'estoqueatual.csv' não encontrado. Exibindo apenas a média de consumo.")
        st.dataframe(media_geral_mensal.style.format({col_media_consumo_nome: "{:,.1f}"}), use_container_width=True)
        excel_data_media = df_to_excel_bytes(media_geral_mensal)
        st.download_button(
            label="📥 Exportar Média de Consumo para Excel",
            data=excel_data_media,
            file_name="media_geral_consumo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        if 'Descricao Requisitante' in material_df.columns and \
           material_df['Descricao Requisitante'].notna().any() and \
           material_df['Descricao Requisitante'].nunique() > 1:
            
            st.header("🏥 Análise de Consumo por Unidade Requisitante")

            st.subheader("Quantidade Total de Consumo por Ano (Agregado por Unidade e Insumo)")
            st.caption("Esta tabela mostra o consumo total para os itens e filtros principais selecionados, distribuído por unidade requisitante. As unidades estão ordenadas pelo seu total consumido.")

            if not analysis_df_materiais.empty:
                df_for_new_table_source = analysis_df_materiais[
                    analysis_df_materiais['Descricao Requisitante'].notna() &
                    (analysis_df_materiais['Descricao Requisitante'] != 'N/A') &
                    (analysis_df_materiais['Descricao Requisitante'].str.strip() != '')
                ].copy()

                if not df_for_new_table_source.empty:
                    try:
                        pivot_unit_consumo_img = pd.pivot_table(
                            df_for_new_table_source,
                            index='Desc. Insumo',
                            columns='Descricao Requisitante',
                            values='Quantidade',
                            aggfunc='sum',
                            fill_value=0
                        )

                        if not pivot_unit_consumo_img.empty:
                            unit_total_consumption_img = pivot_unit_consumo_img.sum(axis=0).sort_values(ascending=False)
                            sorted_unit_columns_img = unit_total_consumption_img.index.tolist()
                            sorted_unit_columns_img = [
                                unit for unit in sorted_unit_columns_img if unit_total_consumption_img[unit] > 0
                            ]

                            if sorted_unit_columns_img:
                                pivot_unit_consumo_sorted_img = pivot_unit_consumo_img[sorted_unit_columns_img]
                                pivot_unit_consumo_sorted_img['Total'] = pivot_unit_consumo_sorted_img[sorted_unit_columns_img].sum(axis=1)
                                total_row_series_img = pivot_unit_consumo_sorted_img.sum(axis=0)
                                total_row_series_img.name = 'Total'
                                display_table_img = pivot_unit_consumo_sorted_img.reset_index()
                                display_table_img.rename(columns={'Desc. Insumo': 'Descricao'}, inplace=True)
                                total_row_df_img = pd.DataFrame(total_row_series_img).T
                                total_row_df_img['Descricao'] = 'Total'
                                current_display_columns = display_table_img.columns.tolist()
                                total_row_df_img = total_row_df_img[current_display_columns]
                                final_display_table_img = pd.concat([display_table_img, total_row_df_img], ignore_index=True)
                                numeric_cols_img = [col for col in final_display_table_img.columns if col != 'Descricao']
                                format_dict_new_table_img = {col: "{:,.0f}" for col in numeric_cols_img}
                                
                                st.dataframe(final_display_table_img.style.format(format_dict_new_table_img, na_rep='0'), use_container_width=True)
                                excel_data_unidade_geral = df_to_excel_bytes(final_display_table_img)
                                st.download_button(
                                    label="📥 Exportar Consumo Agregado por Unidade (Geral) para Excel",
                                    data=excel_data_unidade_geral,
                                    file_name="consumo_agregado_unidade_geral.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                st.info("Nenhuma unidade com consumo significativo encontrado para os itens e filtros selecionados para gerar a tabela de visão geral por unidade.")
                        else:
                            st.info("Não foi possível gerar a tabela de consumo agregado por unidade e insumo (a tabela dinâmica resultou vazia ou sem dados significativos).")
                    except Exception as e:
                        st.error(f"Ocorreu um erro ao gerar a tabela de consumo agregado por unidade: {str(e)}")
                else:
                    st.info("Nenhum dado de consumo por unidade encontrado (após remover N/A ou unidades vazias) para os filtros aplicados para gerar a tabela de visão geral.")
            else:
                st.info("Os dados de análise principal estão vazios. Não é possível gerar a tabela de consumo por unidade.")
            
            st.markdown("---")

            material_para_analise_unidade_global = None
            codigo_insumo_para_unidade_global = ""

            if actual_selected_insumo_descriptions:
                if len(actual_selected_insumo_descriptions) > 1:
                    material_para_analise_unidade_global = st.selectbox(
                        "Selecione UM insumo (por descrição) para analisar consumo detalhado por unidade:",
                        options=actual_selected_insumo_descriptions,
                        index=0,
                        key="select_insumo_unidade_detalhe_v3" 
                    )
                elif len(actual_selected_insumo_descriptions) == 1:
                    material_para_analise_unidade_global = actual_selected_insumo_descriptions[0]
                    st.markdown(f"Analisando consumo detalhado por unidade para o único insumo selecionado: **{material_para_analise_unidade_global}**")
            else:
                 st.info("Nenhum insumo selecionado ou disponível para análise detalhada por unidade.")

            if material_para_analise_unidade_global:
                st.subheader(f"Consumo de '{material_para_analise_unidade_global}' por Unidade (Detalhado)")
                insumo_selecionado_df_para_unidade = material_df[material_df['Desc. Insumo'] == material_para_analise_unidade_global]
                if not insumo_selecionado_df_para_unidade.empty:
                    codigo_insumo_para_unidade_global = str(insumo_selecionado_df_para_unidade['Cód. Insumo'].iloc[0])

                df_unidade_analise_detalhada = analysis_df_materiais[
                    (analysis_df_materiais['Desc. Insumo'] == material_para_analise_unidade_global) &
                    (analysis_df_materiais['Descricao Requisitante'].notna()) &
                    (analysis_df_materiais['Descricao Requisitante'] != 'N/A') &
                    (analysis_df_materiais['Descricao Requisitante'].str.strip() != '')
                ]

                if not df_unidade_analise_detalhada.empty:
                    consumo_total_anual_unidade_df_det = df_unidade_analise_detalhada.groupby(
                        ['Descricao Requisitante', 'Ano']
                    )['Quantidade'].sum().reset_index()

                    meses_com_consumo_unidade_df_det = df_unidade_analise_detalhada[df_unidade_analise_detalhada['Quantidade'] > 0].groupby(
                        ['Descricao Requisitante', 'Ano']
                    )['Mês Num'].nunique().reset_index()
                    meses_com_consumo_unidade_df_det.rename(columns={'Mês Num': 'Nº Meses com Consumo Unidade'}, inplace=True)

                    consumo_unidade_ano_det = pd.merge(
                        consumo_total_anual_unidade_df_det,
                        meses_com_consumo_unidade_df_det,
                        on=['Descricao Requisitante', 'Ano'],
                        how='left'
                    )
                    consumo_unidade_ano_det['Nº Meses com Consumo Unidade'] = consumo_unidade_ano_det['Nº Meses com Consumo Unidade'].fillna(0)
                    
                    consumo_unidade_ano_det['Média Mensal por Unidade'] = consumo_unidade_ano_det.apply(
                        lambda row: row['Quantidade'] / row['Nº Meses com Consumo Unidade'] if row['Nº Meses com Consumo Unidade'] > 0 else 0,
                        axis=1
                    )

                    media_mensal_por_unidade_pdf = consumo_unidade_ano_det.groupby('Descricao Requisitante')['Média Mensal por Unidade'].mean().reset_index().sort_values(by='Média Mensal por Unidade', ascending=False)
                    
                    st.caption(f"Média Mensal de Consumo de '{material_para_analise_unidade_global}' por Unidade (anos {', '.join(map(str,selected_years))}, calculado sobre meses com consumo)");
                    st.dataframe(media_mensal_por_unidade_pdf.style.format({'Média Mensal por Unidade': "{:,.1f}"}), use_container_width=True)
                    if not media_mensal_por_unidade_pdf.empty:
                        excel_data_media_unidade_det = df_to_excel_bytes(media_mensal_por_unidade_pdf)
                        st.download_button(
                            label=f"📥 Exportar Média Mensal ({material_para_analise_unidade_global}) por Unidade para Excel",
                            data=excel_data_media_unidade_det,
                            file_name=f"media_mensal_unidade_{material_para_analise_unidade_global.replace(' ','_').lower()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        fig_unidade_media = px.bar(media_mensal_por_unidade_pdf.head(15), x='Descricao Requisitante', y='Média Mensal por Unidade', color='Descricao Requisitante', title=f'Top 15 Unidades por Média Mensal de Consumo de "{material_para_analise_unidade_global}" (calculado sobre meses com consumo)');
                        st.plotly_chart(fig_unidade_media, use_container_width=True)
                    
                    st.caption(f"Detalhe: Média Mensal por Unidade/Ano para '{material_para_analise_unidade_global}' (calculado sobre meses com consumo)");
                    pivot_unidade_ano_media_mensal_pdf = consumo_unidade_ano_det.pivot_table(index='Descricao Requisitante', columns='Ano', values='Média Mensal por Unidade', fill_value=0).reset_index(); 
                    st.dataframe(pivot_unidade_ano_media_mensal_pdf.style.format({year: "{:,.1f}" for year in selected_years}), height=300, use_container_width=True) 
                    if not pivot_unidade_ano_media_mensal_pdf.empty:
                        excel_data_pivot_unidade_ano = df_to_excel_bytes(pivot_unidade_ano_media_mensal_pdf)
                        st.download_button(
                            label=f"📥 Exportar Detalhe Unidade/Ano ({material_para_analise_unidade_global}) para Excel",
                            data=excel_data_pivot_unidade_ano,
                            file_name=f"detalhe_unidade_ano_{material_para_analise_unidade_global.replace(' ','_').lower()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else: 
                    st.info(f"Nenhum dado de consumo detalhado para '{material_para_analise_unidade_global}' nas unidades e anos selecionados.")
        
        else: 
            if 'Descricao Requisitante' not in material_df.columns:
                 st.warning("A coluna 'Descricao Requisitante' não foi encontrada nos dados. A análise por unidade requisitante não está disponível.")

        if not analysis_df_materiais.empty:
            pdf_bytes = generate_pdf_report(
                actual_selected_insumo_descriptions, selected_cod_insumos, selected_years, 
                selected_movimento_consumo, selected_classes, 
                consumo_anual_pivot_pdf, consumo_mensal_pivot_pdf, 
                fig_consumo_anual_line, fig_consumo_mensal_bar, 
                media_geral_mensal_pdf, 
                consumo_mensal_detalhado_pdf_display,
                fig_consumo_mensal_detalhado,     
                material_para_analise_unidade_global, 
                codigo_insumo_para_unidade_global,  
                media_mensal_por_unidade_pdf, 
                fig_unidade_media, 
                pivot_unidade_ano_media_mensal_pdf
            )
            pdf_download_button_placeholder.download_button(label="📥 Exportar Relatório para PDF", data=pdf_bytes, file_name=f"relatorio_consumo_{'_'.join(map(str,selected_years)) if selected_years else 'geral'}.pdf", mime="application/pdf")
        else: pdf_download_button_placeholder.empty()

st.markdown("---")
st.caption("Dashboard para análise de consumo.")

# --- END OF FILE dashboard_filtro_movimento_indice6.py ---