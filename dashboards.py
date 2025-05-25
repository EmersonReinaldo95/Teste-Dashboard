# --- START OF FILE dashboard_corrigido_completo.py ---

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
    except FileNotFoundError:
        st.error("Arquivo 'Material-CSVANUAL.csv' não encontrado.")
        return pd.DataFrame()
    except pd.errors.EmptyDataError:
        st.error("Arquivo 'Material-CSVANUAL.csv' está vazio.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao ler CSV: {e}")
        return pd.DataFrame()
    return df

# --- Pré-processamento de dados ---
@st.cache_data
def preprocess_data(df_original):
    if df_original.empty:
        return pd.DataFrame()
    df = df_original.copy()

    column_mapping = {
        'Insumo': 'Cód. Insumo',
        'Descricao': 'Desc. Insumo',
        'Dt Movimento': 'Dt Movimento',
        'Quantidade': 'Quantidade',
        'Descricao Movimento': 'Descricao Movimento',
        'Descricao Requisitante': 'Descricao Requisitante',
        'Valor ': 'Valor'
    }
    actual_renames = {}
    # Debugging de colunas lidas (removido para produção final, mas útil durante o desenvolvimento)
    # if not hasattr(preprocess_data, 'logged_csv_columns_mapping_warning'):
    #     st.sidebar.text("DEBUG: Colunas lidas pelo Pandas do CSV:")
    #     st.sidebar.json(df.columns.tolist())
    #     preprocess_data.logged_csv_columns_mapping_warning = True

    for original_name_csv, new_name_internal in column_mapping.items():
        if original_name_csv in df.columns:
            actual_renames[original_name_csv] = new_name_internal
        else:
            st.sidebar.warning(f"Coluna original '{original_name_csv}' do CSV não encontrada para mapeamento. Será ignorada.")
    df.rename(columns=actual_renames, inplace=True)

    if 'Quantidade' not in df.columns: st.error("Coluna interna 'Quantidade' não encontrada após mapeamento."); return pd.DataFrame()
    df['Quantidade'] = pd.to_numeric(df['Quantidade'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
    df['Quantidade'] = df['Quantidade'].abs()

    if 'Dt Movimento' not in df.columns: st.error("Coluna interna 'Dt Movimento' não encontrada após mapeamento."); return pd.DataFrame()
    df['Dt Movimento'] = pd.to_datetime(df['Dt Movimento'], dayfirst=True, errors='coerce')
    df['Ano'] = df['Dt Movimento'].dt.year
    df = df.dropna(subset=['Dt Movimento', 'Ano']); df['Ano'] = df['Ano'].astype(int)

    essential_cols_internal = ['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento', 'Quantidade', 'Ano', 'Dt Movimento']
    optional_cols_internal = ['Descricao Requisitante', 'Valor']
    for col in essential_cols_internal:
        if col not in df.columns:
            st.error(f"Coluna interna essencial '{col}' não encontrada após mapeamento/processamento.");
            st.error(f"DEBUG: Colunas disponíveis no DataFrame neste ponto: {df.columns.tolist()}")
            return pd.DataFrame()
    for col in optional_cols_internal:
        if col not in df.columns:
            st.sidebar.warning(f"Coluna interna opcional '{col}' não encontrada. Será criada como 'N/A' ou 0.")
            df[col] = 0.0 if col == 'Valor' else 'N/A'
    
    df = df.dropna(subset=['Desc. Insumo', 'Cód. Insumo', 'Descricao Movimento'])
    df['Desc. Insumo'] = df['Desc. Insumo'].astype(str); df['Cód. Insumo'] = df['Cód. Insumo'].astype(str)
    if 'Descricao Requisitante' in df.columns: df['Descricao Requisitante'] = df['Descricao Requisitante'].astype(str).fillna('N/A')
    if 'Valor' in df.columns:
        df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
        df['Valor'] = df['Valor'].abs().fillna(0)
    return df

# --- Função para Gerar PDF ---
def generate_pdf_report(
    selected_desc_insumos_pdf, selected_cod_insumos_pdf, selected_years_pdf, selected_movimento_consumo_pdf,
    consumo_anual_pivot_df, consumo_mensal_pivot_df, fig_consumo_anual_line_obj, fig_consumo_mensal_bar_obj,
    media_geral_anual_df, media_geral_mensal_df, material_analise_unidade_pdf=None,
    media_mensal_unidade_df=None, fig_unidade_media_obj=None, pivot_unidade_media_mensal_df=None
    ):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=inch/2, leftMargin=inch/2, topMargin=inch/2, bottomMargin=inch/2)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Relatório de Análise de Consumo de Materiais", styles['h1']))
    story.append(Spacer(1, 0.2*inch))
    if selected_desc_insumos_pdf: story.append(Paragraph(f"<b>Descrições de Insumos Selecionados:</b> {', '.join(selected_desc_insumos_pdf)}", styles['Normal']))
    if selected_cod_insumos_pdf: story.append(Paragraph(f"<b>Códigos de Insumos Selecionados:</b> {', '.join(selected_cod_insumos_pdf)}", styles['Normal']))
    if not selected_desc_insumos_pdf and not selected_cod_insumos_pdf: story.append(Paragraph("<b>Insumos:</b> Nenhum filtro específico aplicado", styles['Normal']))
    story.append(Paragraph(f"<b>Anos Selecionados:</b> {', '.join(map(str, selected_years_pdf)) if selected_years_pdf else 'Nenhum'}", styles['Normal']))
    story.append(Paragraph(f"<b>Tipo de Movimento para Consumo:</b> {selected_movimento_consumo_pdf}", styles['Normal']))
    story.append(Spacer(1, 0.2*inch))

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
        except ValueError as ve: error_msg = f"Erro ao renderizar gráfico '{title}' para PDF (ValueError): {ve}. Verifique Kaleido/dependências."; st.sidebar.error(error_msg); story.append(Paragraph(error_msg, styles['Italic']))
        except Exception as e: error_msg = f"Erro geral ao renderizar gráfico '{title}' para PDF: {e}."; st.sidebar.error(error_msg); story.append(Paragraph(error_msg, styles['Italic']))
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

# Inicializar todas as variáveis que serão passadas para o PDF
fig_consumo_anual_line = None
fig_consumo_mensal_bar = None
fig_media_geral_mensal_grafico = None
fig_unidade_media = None
consumo_anual_pivot_pdf = pd.DataFrame()
consumo_mensal_pivot_pdf = pd.DataFrame()
media_geral_anual_pdf = pd.DataFrame()
media_geral_mensal_pdf = pd.DataFrame()
material_para_analise_unidade_global = None
media_mensal_por_unidade_pdf = pd.DataFrame()
pivot_unidade_ano_media_mensal_pdf = pd.DataFrame()

if material_df.empty:
    st.warning("Não foi possível carregar ou processar os dados. Verifique as mensagens de erro e o mapeamento de colunas na função 'preprocess_data'.")
    st.stop()

st.sidebar.header("⚙️ Filtros de Análise")
all_desc_insumos = sorted(material_df['Desc. Insumo'].dropna().unique())
selected_desc_insumos = st.sidebar.multiselect("💊 Selecione Insumos por Descrição:", options=all_desc_insumos, default=[])
all_cod_insumos = sorted(material_df['Cód. Insumo'].dropna().unique())
selected_cod_insumos = st.sidebar.multiselect("🔢 Selecione Insumos por Código:", options=all_cod_insumos, default=[])
all_years = sorted(material_df['Ano'].dropna().unique())
selected_years = st.sidebar.multiselect("📅 Selecione os Anos:", options=all_years, default=[])
movimento_options = sorted(material_df['Descricao Movimento'].dropna().unique())
default_movimento_str = {movimento_options[6]}
default_movimento_index = 0
if default_movimento_str in movimento_options: default_movimento_index = movimento_options.index(default_movimento_str)
elif movimento_options: st.sidebar.warning(f"Tipo de movimento '{default_movimento_str}' encontrado. Usando '{movimento_options[6]}' como padrão.")
selected_movimento_consumo = st.sidebar.selectbox("📉 Tipo de Movimento para Consumo:", options=movimento_options, index=default_movimento_index if movimento_options else 0)
pdf_download_button_placeholder = st.sidebar.empty()

condition_desc = material_df['Desc. Insumo'].isin(selected_desc_insumos) if selected_desc_insumos else pd.Series(False, index=material_df.index)
condition_cod = material_df['Cód. Insumo'].isin(selected_cod_insumos) if selected_cod_insumos else pd.Series(False, index=material_df.index)
if not selected_desc_insumos and not selected_cod_insumos: df_insumos_selecionados_base = pd.DataFrame(columns=material_df.columns)
else: df_insumos_selecionados_base = material_df[condition_desc | condition_cod]
actual_selected_insumo_descriptions = sorted(df_insumos_selecionados_base['Desc. Insumo'].unique()) if not df_insumos_selecionados_base.empty else []

if not actual_selected_insumo_descriptions: st.info("👈 Por favor, selecione pelo menos um insumo (por descrição ou código) na barra lateral.")
elif not selected_years: st.info("👈 Por favor, selecione pelo menos um ano.")
elif not selected_movimento_consumo: st.info("👈 Por favor, selecione um tipo de movimento.")
else:
    analysis_df_materiais = material_df[(material_df['Desc. Insumo'].isin(actual_selected_insumo_descriptions)) & (material_df['Ano'].isin(selected_years)) & (material_df['Descricao Movimento'] == selected_movimento_consumo)]
    if analysis_df_materiais.empty: st.warning(f"Nenhum dado encontrado para os critérios selecionados.")
    else:
        st.header("🔬 Análise de Consumo por Insumo")
        consumo_anual_por_material = analysis_df_materiais.groupby(['Desc. Insumo', 'Ano'])['Quantidade'].sum().reset_index()
        consumo_anual_por_material.rename(columns={'Quantidade': 'Consumo Total Anual'}, inplace=True)
        consumo_anual_por_material['Consumo Médio Mensal'] = consumo_anual_por_material['Consumo Total Anual'] / 12
        consumo_anual_por_material = consumo_anual_por_material.sort_values(by=['Desc. Insumo', 'Ano'])
        
        st.subheader("Consumo Total Anual")
        try:
            consumo_anual_pivot_pdf = consumo_anual_por_material.pivot_table(index='Desc. Insumo', columns='Ano', values='Consumo Total Anual', fill_value=0).reset_index()
            st.dataframe(consumo_anual_pivot_pdf.style.format({year: "{:,.0f}" for year in selected_years}), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao criar tabela de consumo anual: {e}")
            consumo_anual_pivot_pdf = pd.DataFrame() # Garante que é um DF vazio em caso de erro

        st.subheader("Consumo Médio Mensal")
        try:
            consumo_mensal_pivot_pdf = consumo_anual_por_material.pivot_table(index='Desc. Insumo', columns='Ano', values='Consumo Médio Mensal', fill_value=0).reset_index()
            st.dataframe(consumo_mensal_pivot_pdf.style.format({year: "{:,.1f}" for year in selected_years}), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao criar tabela de consumo mensal: {e}")
            consumo_mensal_pivot_pdf = pd.DataFrame() # Garante que é um DF vazio em caso de erro

        if not consumo_anual_por_material.empty:
            fig_consumo_anual_line = px.line(consumo_anual_por_material, x='Ano', y='Consumo Total Anual', color='Desc. Insumo', markers=True, title='Tendência de Consumo Total Anual por Insumo', labels={'Desc. Insumo': 'Insumo'})
            st.plotly_chart(fig_consumo_anual_line.update_layout(xaxis_type='category'), use_container_width=True)
            fig_consumo_mensal_bar = px.bar(consumo_anual_por_material, x='Ano', y='Consumo Médio Mensal', color='Desc. Insumo', barmode='group', title='Comparativo de Consumo Médio Mensal por Insumo', labels={'Desc. Insumo': 'Insumo'})
            st.plotly_chart(fig_consumo_mensal_bar.update_layout(xaxis_type='category'), use_container_width=True)
        
        if len(selected_years) > 0:
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
                df_unidade_analise = material_df[(material_df['Desc. Insumo'] == material_para_analise_unidade_global) & (material_df['Ano'].isin(selected_years)) & (material_df['Descricao Movimento'] == selected_movimento_consumo) & (material_df['Descricao Requisitante'].notna()) & (material_df['Descricao Requisitante'] != 'N/A')]
                if not df_unidade_analise.empty:
                    consumo_unidade_ano = df_unidade_analise.groupby(['Descricao Requisitante', 'Ano'])['Quantidade'].sum().reset_index(); consumo_unidade_ano['Média Mensal por Unidade'] = consumo_unidade_ano['Quantidade'] / 12
                    media_mensal_por_unidade_pdf = consumo_unidade_ano.groupby('Descricao Requisitante')['Média Mensal por Unidade'].mean().reset_index().sort_values(by='Média Mensal por Unidade', ascending=False)
                    st.caption(f"Média Mensal de Consumo de '{material_para_analise_unidade_global}' por Unidade (anos {', '.join(map(str,selected_years))})"); st.dataframe(media_mensal_por_unidade_pdf.style.format({'Média Mensal por Unidade': "{:,.1f}"}), use_container_width=True)
                    if not media_mensal_por_unidade_pdf.empty: fig_unidade_media = px.bar(media_mensal_por_unidade_pdf.head(15), x='Descricao Requisitante', y='Média Mensal por Unidade', color='Descricao Requisitante', title=f'Top 15 Unidades por Média Mensal de Consumo de "{material_para_analise_unidade_global}"'); st.plotly_chart(fig_unidade_media, use_container_width=True)
                    st.caption(f"Detalhe: Média Mensal por Unidade/Ano para '{material_para_analise_unidade_global}'"); pivot_unidade_ano_media_mensal_pdf = consumo_unidade_ano.pivot_table(index='Descricao Requisitante', columns='Ano', values='Média Mensal por Unidade', fill_value=0).reset_index(); st.dataframe(pivot_unidade_ano_media_mensal_pdf.style.format({year: "{:,.1f}" for year in selected_years}), height=300, use_container_width=True)
                else: st.info(f"Nenhum dado de consumo para '{material_para_analise_unidade_global}' nas unidades/anos selecionados.")
        else: st.info("Análise por unidade desabilitada (coluna 'Descricao Requisitante' ausente/inválida ou com poucas unidades).")
        
        pdf_bytes = generate_pdf_report(actual_selected_insumo_descriptions, selected_cod_insumos, selected_years, selected_movimento_consumo, consumo_anual_pivot_pdf, consumo_mensal_pivot_pdf, fig_consumo_anual_line, fig_consumo_mensal_bar, media_geral_anual_pdf, media_geral_mensal_pdf, material_para_analise_unidade_global, media_mensal_por_unidade_pdf, fig_unidade_media, pivot_unidade_ano_media_mensal_pdf)
        pdf_download_button_placeholder.download_button(label="📥 Exportar Relatório para PDF", data=pdf_bytes, file_name=f"relatorio_consumo_{'_'.join(map(str,selected_years)) if selected_years else 'geral'}.pdf", mime="application/pdf")

st.markdown("---")
st.caption("Dashboard para análise de consumo.")

# --- END OF FILE dashboard_corrigido_completo.py ---
