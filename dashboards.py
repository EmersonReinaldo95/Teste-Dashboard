# --- START OF FILE dashboard_filtro_movimento_indice6.py ---

import streamlit as st
import pandas as pd
import plotly.express as px
import io
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
import plotly.io as pio

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
    df = df.dropna(subset=['Dt Movimento', 'Ano']); df['Ano'] = df['Ano'].astype(int)

    essential_cols_internal = ['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento', 'Quantidade', 'Ano', 'Dt Movimento']
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

# --- Função para Gerar PDF (sem alterações nesta função) ---
def generate_pdf_report(
    selected_desc_insumos_pdf, selected_cod_insumos_pdf, selected_years_pdf,
    selected_movimento_consumo_pdf, selected_classes_pdf,
    consumo_anual_pivot_df, consumo_mensal_pivot_df, fig_consumo_anual_line_obj, fig_consumo_mensal_bar_obj,
    media_geral_anual_df, media_geral_mensal_df, material_analise_unidade_pdf=None,
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
        if title: story.append(Paragraph(f"<b>{title}</b>", styles['h3'])); story.append(Spacer(1, 0.1*inch))
        if df.empty: story.append(Paragraph("Nenhum dado para exibir.", styles['Italic'])); story.append(Spacer(1, 0.1*inch)); return
        max_cols = 10; df_display = df.iloc[:, :max_cols].copy() if len(df.columns) > max_cols else df.copy()
        if len(df.columns) > max_cols: df_display['...'] = '...' ; story.append(Paragraph(f"(Exibindo as primeiras {max_cols-1} colunas de dados e coluna de índice)", styles['Italic']))
        data = [df_display.columns.to_list()] + df_display.astype(str).values.tolist()
        table = Table(data, repeatRows=1); table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),('FONTSIZE', (0,0), (-1,-1), 7), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ])); story.append(table); story.append(Spacer(1, 0.2*inch))

    def fig_to_image_reportlab(fig, title=""):
        if title: story.append(Paragraph(f"<b>{title}</b>", styles['h3'])); story.append(Spacer(1, 0.1*inch))
        if fig is None: story.append(Paragraph("Gráfico não disponível.", styles['Italic'])); story.append(Spacer(1, 0.1*inch)); return
        try:
            img_bytes = pio.to_image(fig, format="png", width=700, height=350, scale=1.5)
            img = Image(io.BytesIO(img_bytes), width=6.8*inch, height=3.4*inch)
            img.hAlign = 'CENTER'; story.append(img); story.append(Spacer(1, 0.2*inch))
        except Exception as e: error_msg = f"Erro renderizar gráfico '{title}' PDF: {e}"; st.sidebar.error(error_msg); story.append(Paragraph(error_msg, styles['Italic']))
        story.append(Spacer(1, 0.1*inch))
    
    story.append(Paragraph("Análise de Consumo por Insumo", styles['h2']))
    df_to_table(consumo_anual_pivot_df, "Consumo Total Anual por Insumo (Agrupado por Desc. Insumo)")
    df_to_table(consumo_mensal_pivot_df, "Consumo Médio Mensal por Insumo (Agrupado por Desc. Insumo)")
    fig_to_image_reportlab(fig_consumo_anual_line_obj, "Tendência de Consumo Total Anual")
    fig_to_image_reportlab(fig_consumo_mensal_bar_obj, "Comparativo de Consumo Médio Mensal")
    story.append(Paragraph("Médias Gerais de Consumo por Insumo (sobre os anos selecionados)", styles['h2']))
    df_to_table(media_geral_anual_df, "Média Geral Anual por Insumo")
    df_to_table(media_geral_mensal_df, "Média Geral Mensal por Insumo")
    if material_analise_unidade_pdf and media_mensal_unidade_df is not None and not media_mensal_unidade_df.empty:
        story.append(PageBreak())
        story.append(Paragraph(f"Análise de Consumo por Unidade para o Insumo: {material_analise_unidade_pdf}", styles['h2']))
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
consumo_anual_pivot_pdf = pd.DataFrame(); consumo_mensal_pivot_pdf = pd.DataFrame()
media_geral_anual_pdf = pd.DataFrame(); media_geral_mensal_pdf = pd.DataFrame()
material_para_analise_unidade_global = None
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

# --- FILTRO DE TIPO DE MOVIMENTO COM DEFAULT ESPECÍFICO (ÍNDICE 6) ---
movimento_options = sorted(material_df['Descricao Movimento'].dropna().unique())
default_movimento_index = 0 # Padrão para o primeiro item

if len(movimento_options) > 6: # Verifica se há pelo menos 7 itens (para que o índice 6 seja válido)
    default_movimento_index = 6
    # Opcional: informar qual foi o default se não for o primeiro
    # st.sidebar.info(f"Tipo de Movimento padrão: '{movimento_options[default_movimento_index]}'")
elif movimento_options: # Se não há 7 itens, mas há opções, usa o primeiro
    default_movimento_index = 0
    st.sidebar.warning(f"Menos de 7 tipos de movimento disponíveis. Usando '{movimento_options[0]}' como padrão.")
# Se movimento_options estiver vazio, o selectbox não terá opções e o index 0 não causará erro.

selected_movimento_consumo = st.sidebar.selectbox(
    "📉 Tipo de Movimento para Consumo:",
    options=movimento_options,
    index=default_movimento_index if movimento_options else 0 # Proteção para lista vazia
)
# --- FIM DO FILTRO DE TIPO DE MOVIMENTO ---

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
        # --- CORPO DO DASHBOARD (TABELAS, GRÁFICOS, ANÁLISE POR UNIDADE) ---
        # Esta parte do código permanece a mesma da versão anterior,
        # começando com st.header("🔬 Análise de Consumo por Insumo")
        # e terminando antes da última chamada a st.markdown("---")
        st.header("🔬 Análise de Consumo por Insumo")
        consumo_anual_por_material = analysis_df_materiais.groupby(['Desc. Insumo', 'Ano'])['Quantidade'].sum().reset_index()
        consumo_anual_por_material.rename(columns={'Quantidade': 'Consumo Total Anual'}, inplace=True)
        consumo_anual_por_material['Consumo Médio Mensal'] = consumo_anual_por_material['Consumo Total Anual'] / 12
        consumo_anual_por_material = consumo_anual_por_material.sort_values(by=['Desc. Insumo', 'Ano'])
        
        st.subheader("Consumo Total Anual")
        try:
            consumo_anual_pivot_pdf = consumo_anual_por_material.pivot_table(index='Desc. Insumo', columns='Ano', values='Consumo Total Anual', fill_value=0).reset_index()
            st.dataframe(consumo_anual_pivot_pdf.style.format({year: "{:,.0f}" for year in selected_years}), use_container_width=True)
        except Exception as e: st.error(f"Erro ao criar tabela de consumo anual: {e}"); consumo_anual_pivot_pdf = pd.DataFrame()

        st.subheader("Consumo Médio Mensal")
        try:
            consumo_mensal_pivot_pdf = consumo_anual_por_material.pivot_table(index='Desc. Insumo', columns='Ano', values='Consumo Médio Mensal', fill_value=0).reset_index()
            st.dataframe(consumo_mensal_pivot_pdf.style.format({year: "{:,.1f}" for year in selected_years}), use_container_width=True)
        except Exception as e: st.error(f"Erro ao criar tabela de consumo mensal: {e}"); consumo_mensal_pivot_pdf = pd.DataFrame()

        if not consumo_anual_por_material.empty:
            fig_consumo_anual_line = px.line(consumo_anual_por_material, x='Ano', y='Consumo Total Anual', color='Desc. Insumo', markers=True, title='Tendência de Consumo Total Anual por Insumo', labels={'Desc. Insumo': 'Insumo'}); st.plotly_chart(fig_consumo_anual_line.update_layout(xaxis_type='category'), use_container_width=True)
            fig_consumo_mensal_bar = px.bar(consumo_anual_por_material, x='Ano', y='Consumo Médio Mensal', color='Desc. Insumo', barmode='group', title='Comparativo de Consumo Médio Mensal por Insumo', labels={'Desc. Insumo': 'Insumo'}); st.plotly_chart(fig_consumo_mensal_bar.update_layout(xaxis_type='category'), use_container_width=True)
        
        if len(selected_years) > 0 and not consumo_anual_por_material.empty :
            media_geral_anual_pdf = consumo_anual_por_material.groupby('Desc. Insumo')['Consumo Total Anual'].mean().reset_index(); media_geral_anual_pdf.rename(columns={'Consumo Total Anual': f'Média Geral Anual ({len(selected_years)}a)'}, inplace=True)
            media_geral_mensal_pdf = consumo_anual_por_material.groupby('Desc. Insumo')['Consumo Médio Mensal'].mean().reset_index(); media_geral_mensal_pdf.rename(columns={'Consumo Médio Mensal': f'Média Geral Mensal ({len(selected_years)}a)'}, inplace=True)
            st.subheader(f"⚖️ Médias Gerais de Consumo por Insumo (sobre Anos Selecionados)"); col_media1, col_media2 = st.columns(2)
            with col_media1: st.caption("Média Geral Anual"); st.dataframe(media_geral_anual_pdf.style.format({media_geral_anual_pdf.columns[1]: "{:,.0f}"}), use_container_width=True)
            with col_media2: st.caption("Média Geral Mensal"); st.dataframe(media_geral_mensal_pdf.style.format({media_geral_mensal_pdf.columns[1]: "{:,.1f}"}), use_container_width=True)
            if not media_geral_mensal_pdf.empty: fig_media_geral_mensal_grafico = px.bar(media_geral_mensal_pdf, x='Desc. Insumo', y=media_geral_mensal_pdf.columns[1], color='Desc. Insumo', title='Média Geral do Consumo Mensal por Insumo', labels={'Desc. Insumo': 'Insumo'}); st.plotly_chart(fig_media_geral_mensal_grafico, use_container_width=True)
        
        st.markdown("---")
        if 'Descricao Requisitante' in material_df.columns and material_df['Descricao Requisitante'].notna().any() and material_df['Descricao Requisitante'].nunique() > 1 :
            st.header("🏥 Análise de Consumo por Unidade Requisitante"); material_para_analise_unidade_global = None
            if len(actual_selected_insumo_descriptions) > 1: material_para_analise_unidade_global = st.selectbox("Selecione UM insumo (por descrição) para analisar consumo por unidade:", options=actual_selected_insumo_descriptions, index=0)
            elif len(actual_selected_insumo_descriptions) == 1: material_para_analise_unidade_global = actual_selected_insumo_descriptions[0]
            if material_para_analise_unidade_global:
                st.subheader(f"Consumo de '{material_para_analise_unidade_global}' por Unidade")
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
            pdf_bytes = generate_pdf_report(actual_selected_insumo_descriptions, selected_cod_insumos, selected_years, selected_movimento_consumo, selected_classes, consumo_anual_pivot_pdf, consumo_mensal_pivot_pdf, fig_consumo_anual_line, fig_consumo_mensal_bar, media_geral_anual_pdf, media_geral_mensal_pdf, material_para_analise_unidade_global, media_mensal_por_unidade_pdf, fig_unidade_media, pivot_unidade_ano_media_mensal_pdf)
            pdf_download_button_placeholder.download_button(label="📥 Exportar Relatório para PDF", data=pdf_bytes, file_name=f"relatorio_consumo_{'_'.join(map(str,selected_years)) if selected_years else 'geral'}.pdf", mime="application/pdf")
        else: pdf_download_button_placeholder.empty()

st.markdown("---")
st.caption("Dashboard para análise de consumo.")

# --- END OF FILE dashboard_filtro_movimento_indice6.py ---
