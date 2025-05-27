# --- START OF FILE dashboard_filtro_movimento_indice6.py ---

import streamlit as st
import pandas as pd
import plotly.express as px
import io
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
import plotly.io as pio
# import calendar # Removido
# import locale # Removido, a menos que seja usado para outros fins de formatação de data/hora

# Lista explícita de meses em português para ordenação e exibição correta
MESES_PT_ORDENADOS = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
MESES_PT_MAP = {i+1: mes for i, mes in enumerate(MESES_PT_ORDENADOS)}


# Configuração da página
st.set_page_config(layout="wide", page_title="Dashboard de Consumo com PDF")

# --- Funções de Carregamento ---
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

# --- Pré-processamento de dados ---
@st.cache_data
def preprocess_data(df_original):
    if df_original.empty: return pd.DataFrame()
    df = df_original.copy()
    column_mapping = {
        'Insumo': 'Cód. Insumo', 'Descricao': 'Desc. Insumo', 'Dt Movimento': 'Dt Movimento',
        'Quantidade': 'Quantidade', 'Descricao Movimento': 'Descricao Movimento',
        'Descricao Requisitante': 'Descricao Requisitante', 'Valor ': 'Valor',
        'Descricao da Classe': 'Descricao Classe'
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
    
    # Usar o mapeamento manual para Mês Nome
    df['Mês Nome'] = df['Mês Num'].map(MESES_PT_MAP)
    
    df = df.dropna(subset=['Dt Movimento', 'Ano', 'Mês Num', 'Mês Nome']); 
    df['Ano'] = df['Ano'].astype(int)
    df['Mês Num'] = df['Mês Num'].astype(int)


    essential_cols_internal = ['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento', 'Quantidade', 'Ano', 'Dt Movimento', 'Mês Num', 'Mês Nome']
    optional_cols_internal = ['Descricao Requisitante', 'Valor', 'Descricao Classe']
    for col in essential_cols_internal:
        if col not in df.columns: st.error(f"Coluna interna essencial '{col}' não encontrada."); return pd.DataFrame()
    for col in optional_cols_internal:
        if col not in df.columns: df[col] = 0.0 if col == 'Valor' else 'N/A'
    
    df = df.dropna(subset=['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento'])
    df['Desc. Insumo'] = df['Desc. Insumo'].astype(str); df['Cód. Insumo'] = df['Cód. Insumo'].astype(str)
    if 'Descricao Requisitante' in df.columns: df['Descricao Requisitante'] = df['Descricao Requisitante'].astype(str).fillna('N/A')
    if 'Descricao Classe' in df.columns: df['Descricao Classe'] = df['Descricao Classe'].astype(str).fillna('N/A')
    if 'Valor' in df.columns: df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').abs().fillna(0)
    return df

# --- Função para Gerar PDF (sem alterações nesta função em relação à última versão, exceto que o código acima já prepara os dados corretamente) ---
def generate_pdf_report(
    selected_desc_insumos_pdf, selected_cod_insumos_pdf, selected_years_pdf,
    selected_movimento_consumo_pdf, selected_classes_pdf,
    consumo_anual_pivot_df, consumo_mensal_pivot_df, fig_consumo_anual_line_obj, fig_consumo_mensal_bar_obj,
    media_geral_anual_df, media_geral_mensal_df,
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
        
        max_cols = 10 
        df_display = df.copy()
        if isinstance(df_display.columns, pd.MultiIndex):
            df_display.columns = ['_'.join(map(str, col)).strip('_') for col in df_display.columns.values]

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
            error_msg = f"Erro renderizar gráfico '{title}' PDF: {e}"
            st.sidebar.error(error_msg)
            elements_for_keeptogether.append(Paragraph(error_msg, styles['Italic']))
        
        elements_for_keeptogether.append(Spacer(1, 0.1*inch)) 
        story.append(KeepTogether(elements_for_keeptogether))
    
    story.append(Paragraph("Análise de Consumo por Insumo", styles['h2']))
    df_to_table(consumo_anual_pivot_df, "Consumo Total Anual por Insumo")
    df_to_table(consumo_mensal_pivot_df, "Consumo Médio Mensal Agregado por Ano") 
    fig_to_image_reportlab(fig_consumo_anual_line_obj, "Tendência de Consumo Total Anual")
    fig_to_image_reportlab(fig_consumo_mensal_bar_obj, "Comparativo de Consumo Médio Mensal Agregado por Ano")

    story.append(PageBreak()) 
    story.append(Paragraph("Análise Detalhada de Consumo Mensal", styles['h2']))
    if consumo_mensal_detalhado_pdf_data is not None and not consumo_mensal_detalhado_pdf_data.empty:
        df_to_table(consumo_mensal_detalhado_pdf_data, "Consumo Mensal Efetivo por Insumo (Média entre Anos se >1 selecionado)")
    else:
        story.append(Paragraph("Dados de consumo mensal detalhado não disponíveis.", styles['Normal']))
    fig_to_image_reportlab(fig_consumo_mensal_detalhado_obj, "Gráfico de Consumo Mensal Efetivo por Insumo")


    story.append(PageBreak()) 
    story.append(Paragraph("Médias Gerais de Consumo por Insumo (sobre os anos selecionados)", styles['h2']))
    df_to_table(media_geral_anual_df, "Média Geral Anual por Insumo")
    df_to_table(media_geral_mensal_df, "Média Geral Mensal por Insumo")

    if material_analise_unidade_pdf and media_mensal_unidade_df is not None and not media_mensal_unidade_df.empty:
        story.append(PageBreak())
        titulo_unidade = f"Análise de Consumo por Unidade para o Insumo: {material_analise_unidade_pdf}"
        if material_analise_unidade_cod_insumo_pdf:
            titulo_unidade += f" (Cód: {material_analise_unidade_cod_insumo_pdf})"
        story.append(Paragraph(titulo_unidade, styles['h2']))
        
        df_to_table(media_mensal_unidade_df, f"Média Mensal de Consumo por Unidade")
        if fig_unidade_media_obj: fig_to_image_reportlab(fig_unidade_media_obj, f"Top Unidades por Média Mensal de Consumo")
        if pivot_unidade_media_mensal_df is not None and not pivot_unidade_media_mensal_df.empty: df_to_table(pivot_unidade_media_mensal_df, "Detalhe: Média Mensal por Unidade/Ano")
    doc.build(story); buffer.seek(0); return buffer.getvalue()

# --- Carregar e pré-processar os dados ---
raw_df = load_data()
material_df = preprocess_data(raw_df)

# --- Interface do Dashboard ---
st.title("📊 Dashboard Avançado de Análise de Consumo")

fig_consumo_anual_line = None; fig_consumo_mensal_bar = None; fig_media_geral_mensal_grafico = None; fig_unidade_media = None
fig_consumo_mensal_detalhado = None 
consumo_anual_pivot_pdf = pd.DataFrame(); consumo_mensal_pivot_pdf = pd.DataFrame()
media_geral_anual_pdf = pd.DataFrame(); media_geral_mensal_pdf = pd.DataFrame()
consumo_mensal_detalhado_pdf = pd.DataFrame() 
material_para_analise_unidade_global = None
codigo_insumo_para_unidade_global = "" 
media_mensal_por_unidade_pdf = pd.DataFrame(); pivot_unidade_ano_media_mensal_pdf = pd.DataFrame()

if material_df.empty: st.warning("Dados não carregados/processados. Verifique mapeamento de colunas."); st.stop()

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
elif movimento_options: default_movimento_index = 0; st.sidebar.warning(f"Menos de 7 tipos de movimento disponíveis. Usando '{movimento_options[0]}' como padrão.")
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
elif not selected_movimento_consumo and not movimento_options: st.error("Nenhum tipo de movimento disponível nos dados."); proceed_with_analysis = False
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
        consumo_anual_por_material = analysis_df_materiais.groupby(['Cód. Insumo', 'Desc. Insumo', 'Ano'])['Quantidade'].sum().reset_index()
        consumo_anual_por_material.rename(columns={'Quantidade': 'Consumo Total Anual'}, inplace=True)
        consumo_anual_por_material['Consumo Médio Mensal (agregado)'] = consumo_anual_por_material['Consumo Total Anual'] / 12 
        consumo_anual_por_material = consumo_anual_por_material.sort_values(by=['Cód. Insumo', 'Desc. Insumo', 'Ano'])
        
        st.subheader("Consumo Total Anual")
        try:
            consumo_anual_pivot_pdf = consumo_anual_por_material.pivot_table(index=['Cód. Insumo', 'Desc. Insumo'], columns='Ano', values='Consumo Total Anual', fill_value=0).reset_index()
            st.dataframe(consumo_anual_pivot_pdf.style.format({year: "{:,.0f}" for year in selected_years}), use_container_width=True)
        except Exception as e: st.error(f"Erro ao criar tabela de consumo anual: {e}"); consumo_anual_pivot_pdf = pd.DataFrame()

        st.subheader("Consumo Médio Mensal (agregado por ano)")
        try:
            consumo_mensal_pivot_pdf = consumo_anual_por_material.pivot_table(index=['Cód. Insumo', 'Desc. Insumo'], columns='Ano', values='Consumo Médio Mensal (agregado)', fill_value=0).reset_index()
            st.dataframe(consumo_mensal_pivot_pdf.style.format({year: "{:,.1f}" for year in selected_years}), use_container_width=True)
        except Exception as e: st.error(f"Erro ao criar tabela de consumo mensal agregada: {e}"); consumo_mensal_pivot_pdf = pd.DataFrame()

        if not consumo_anual_por_material.empty:
            fig_consumo_anual_line = px.line(consumo_anual_por_material, x='Ano', y='Consumo Total Anual', color='Desc. Insumo', markers=True, title='Tendência de Consumo Total Anual por Insumo', labels={'Desc. Insumo': 'Insumo'}, hover_data=['Cód. Insumo']); st.plotly_chart(fig_consumo_anual_line.update_layout(xaxis_type='category'), use_container_width=True)
            fig_consumo_mensal_bar = px.bar(consumo_anual_por_material, x='Ano', y='Consumo Médio Mensal (agregado)', color='Desc. Insumo', barmode='group', title='Comparativo de Consumo Médio Mensal (agregado por ano)', labels={'Desc. Insumo': 'Insumo'}, hover_data=['Cód. Insumo']); st.plotly_chart(fig_consumo_mensal_bar.update_layout(xaxis_type='category'), use_container_width=True)

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
        
        consumo_mensal_detalhado_pdf = consumo_mensal_grafico_df.rename(columns={'Quantidade': 'Consumo Mensal Efetivo'})
        if 'Mês Num' in consumo_mensal_detalhado_pdf.columns:
            consumo_mensal_detalhado_pdf_display = consumo_mensal_detalhado_pdf.drop(columns=['Mês Num'])
        else:
            consumo_mensal_detalhado_pdf_display = consumo_mensal_detalhado_pdf
        
        st.subheader("Tabela: " + titulo_detalhado)
        st.dataframe(consumo_mensal_detalhado_pdf_display.style.format({'Consumo Mensal Efetivo': "{:,.1f}"}), use_container_width=True)

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
                # Usar a lista explícita de meses para category_orders
                category_orders={"Mês Nome": MESES_PT_ORDENADOS} 
            )
            st.plotly_chart(fig_consumo_mensal_detalhado, use_container_width=True)
        else:
            fig_consumo_mensal_detalhado = None 
        
        st.markdown("---")
        if len(selected_years) > 0 and not consumo_anual_por_material.empty :
            media_geral_anual_pdf = consumo_anual_por_material.groupby(['Cód. Insumo', 'Desc. Insumo'])['Consumo Total Anual'].mean().reset_index(); media_geral_anual_pdf.rename(columns={'Consumo Total Anual': f'Média Geral Anual ({len(selected_years)}a)'}, inplace=True)
            media_geral_mensal_pdf = consumo_anual_por_material.groupby(['Cód. Insumo', 'Desc. Insumo'])['Consumo Médio Mensal (agregado)'].mean().reset_index(); media_geral_mensal_pdf.rename(columns={'Consumo Médio Mensal (agregado)': f'Média Geral Mensal ({len(selected_years)}a)'}, inplace=True)
            
            st.subheader(f"⚖️ Médias Gerais de Consumo por Insumo (sobre Anos Selecionados)"); col_media1, col_media2 = st.columns(2)
            
            col_valor_anual_idx = 2 
            col_valor_mensal_idx = 2 

            with col_media1: 
                st.caption("Média Geral Anual"); 
                if not media_geral_anual_pdf.empty: st.dataframe(media_geral_anual_pdf.style.format({media_geral_anual_pdf.columns[col_valor_anual_idx]: "{:,.0f}"}), use_container_width=True)
            with col_media2: 
                st.caption("Média Geral Mensal (baseada na média anual / 12)"); 
                if not media_geral_mensal_pdf.empty: st.dataframe(media_geral_mensal_pdf.style.format({media_geral_mensal_pdf.columns[col_valor_mensal_idx]: "{:,.1f}"}), use_container_width=True)

            if not media_geral_mensal_pdf.empty:
                col_valor_media_mensal_nome = media_geral_mensal_pdf.columns[col_valor_mensal_idx]
                fig_media_geral_mensal_grafico = px.bar(media_geral_mensal_pdf, x='Desc. Insumo', y=col_valor_media_mensal_nome, color='Desc. Insumo', title='Média Geral do Consumo Mensal por Insumo', labels={'Desc. Insumo': 'Insumo', col_valor_media_mensal_nome: 'Média Mensal'}, hover_data=['Cód. Insumo'])
                st.plotly_chart(fig_media_geral_mensal_grafico, use_container_width=True)
        
        st.markdown("---")
        if 'Descricao Requisitante' in material_df.columns and material_df['Descricao Requisitante'].notna().any() and material_df['Descricao Requisitante'].nunique() > 1 :
            st.header("🏥 Análise de Consumo por Unidade Requisitante"); 
            material_para_analise_unidade_global = None 
            codigo_insumo_para_unidade_global = ""    

            if len(actual_selected_insumo_descriptions) > 1: 
                material_para_analise_unidade_global = st.selectbox("Selecione UM insumo (por descrição) para analisar consumo por unidade:", options=actual_selected_insumo_descriptions, index=0)
            elif len(actual_selected_insumo_descriptions) == 1: 
                material_para_analise_unidade_global = actual_selected_insumo_descriptions[0]
            
            if material_para_analise_unidade_global:
                st.subheader(f"Consumo de '{material_para_analise_unidade_global}' por Unidade")
                insumo_selecionado_df = material_df[material_df['Desc. Insumo'] == material_para_analise_unidade_global]
                if not insumo_selecionado_df.empty:
                    codigo_insumo_para_unidade_global = str(insumo_selecionado_df['Cód. Insumo'].iloc[0])

                df_unidade_analise = analysis_df_materiais[(analysis_df_materiais['Desc. Insumo'] == material_para_analise_unidade_global) & (analysis_df_materiais['Descricao Requisitante'].notna()) & (analysis_df_materiais['Descricao Requisitante'] != 'N/A')]
                if not df_unidade_analise.empty:
                    consumo_unidade_ano = df_unidade_analise.groupby(['Descricao Requisitante', 'Ano'])['Quantidade'].sum().reset_index(); consumo_unidade_ano['Média Mensal por Unidade'] = consumo_unidade_ano['Quantidade'] / 12
                    media_mensal_por_unidade_pdf = consumo_unidade_ano.groupby('Descricao Requisitante')['Média Mensal por Unidade'].mean().reset_index().sort_values(by='Média Mensal por Unidade', ascending=False)
                    st.caption(f"Média Mensal de Consumo de '{material_para_analise_unidade_global}' por Unidade (anos {', '.join(map(str,selected_years))})"); st.dataframe(media_mensal_por_unidade_pdf.style.format({'Média Mensal por Unidade': "{:,.1f}"}), use_container_width=True)
                    if not media_mensal_por_unidade_pdf.empty: fig_unidade_media = px.bar(media_mensal_por_unidade_pdf.head(15), x='Descricao Requisitante', y='Média Mensal por Unidade', color='Descricao Requisitante', title=f'Top 15 Unidades por Média Mensal de Consumo de "{material_para_analise_unidade_global}"'); st.plotly_chart(fig_unidade_media, use_container_width=True)
                    st.caption(f"Detalhe: Média Mensal por Unidade/Ano para '{material_para_analise_unidade_global}'"); pivot_unidade_ano_media_mensal_pdf = consumo_unidade_ano.pivot_table(index='Descricao Requisitante', columns='Ano', values='Média Mensal por Unidade', fill_value=0).reset_index(); st.dataframe(pivot_unidade_ano_media_mensal_pdf.style.format({year: "{:,.1f}" for year in selected_years}), height=300, use_container_width=True)
                else: st.info(f"Nenhum dado de consumo para '{material_para_analise_unidade_global}' nas unidades e anos selecionados.")
        else: st.info("Análise por unidade desabilitada (coluna 'Descricao Requisitante' ausente/inválida ou com poucas unidades).")
        
        if not analysis_df_materiais.empty:
            pdf_bytes = generate_pdf_report(
                actual_selected_insumo_descriptions, selected_cod_insumos, selected_years, 
                selected_movimento_consumo, selected_classes, 
                consumo_anual_pivot_pdf, consumo_mensal_pivot_pdf, 
                fig_consumo_anual_line, fig_consumo_mensal_bar, 
                media_geral_anual_pdf, media_geral_mensal_pdf, 
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
