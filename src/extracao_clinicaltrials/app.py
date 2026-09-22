# Bibliotecas
from __future__ import annotations

import io
import json
import openpyxl
from io import BytesIO

import altair as alt
import pandas as pd
import streamlit as st
import vl_convert as vlc
from PIL import Image as PILImage
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer

from client import (
    ClinicalTrialsClient,
    DEFAULT_FIELDS,
    load_sponsors_from_file,
    load_values_from_file,
    run_etl,
    run_multi_value_batch,
    run_sponsor_batch,
)

st.set_page_config(page_title="ClinicalTrials.gov — Extrator", page_icon="🧪", layout="wide")

st.title("🧪 ClinicalTrials.gov — Extrator de Dados (API v2)")

# ---------------------------------------------------------------------------
# Verificar versão da API
# ---------------------------------------------------------------------------

@st.cache_data(ttl=3600)
def _get_api_version() -> dict[str, str]:
    client = ClinicalTrialsClient()
    return client.get_version()


# try:
#     version = _get_api_version()
#     st.success(f"API v{version.get('apiVersion')}  |  Dados atualizados em: {version.get('dataTimestamp')}")
# except Exception as e:
#     st.warning(f"Não foi possível verificar a versão da API: {e}")

try:
    version = _get_api_version()
    raw_ts = version.get('dataTimestamp', '')
    timestamp = raw_ts.replace('T', ' ')
    st.success(f"API v{version.get('apiVersion')} | Dados atualizados em: {timestamp}")
except Exception as e:
    st.warning(f"Não foi possível verificar a versão da API: {e}")


def _essie_clause(field: str, values: list[str]) -> str | None:
    """Monta uma cláusula Essie AREA[field]valor(es) a partir de uma lista de seleção."""
    if not values:
        return None
    if len(values) == 1:
        return f"AREA[{field}]{values[0]}"
    return f"AREA[{field}]({' OR '.join(values)})"


def _enrollment_clause(min_value: float, max_value: float) -> str | None:
    """Monta a cláusula Essie de RANGE para EnrollmentCount."""
    if not min_value and not max_value:
        return None
    lo = str(int(min_value)) if min_value else "MIN"
    hi = str(int(max_value)) if max_value else "MAX"
    return f"AREA[EnrollmentCount]RANGE[{lo},{hi}]"


def _dedupe_by_consulted_values(df: pd.DataFrame) -> pd.DataFrame:
    """Une estudos duplicados (mesmo nct_id) resultantes de buscas em lote, concatenando os valores consultados."""
    consulted_columns = [c for c in df.columns if c.endswith("_consultada") or c.endswith("_consultado")]
    if not consulted_columns or "nct_id" not in df.columns or not df["nct_id"].duplicated().any():
        return df

    def _join_unique(values: pd.Series) -> str:
        return "; ".join(dict.fromkeys(v for v in values if v))

    other_columns = [c for c in df.columns if c not in consulted_columns and c != "nct_id"]
    aggregations = {col: _join_unique for col in consulted_columns}
    aggregations.update({col: "first" for col in other_columns})
    return df.groupby("nct_id", as_index=False, sort=False).agg(aggregations)


def _chart_to_png(chart: alt.Chart) -> bytes:
    """Rasteriza um gráfico Altair/Vega-Lite em PNG via vl-convert, com tamanho fixo e consistente."""
    sized_chart = chart.properties(width=900, height=540)
    return vlc.vegalite_to_png(vl_spec=sized_chart.to_dict(), scale=2)


def _fit_box(png_bytes: bytes, max_width_in: float, max_height_in: float) -> tuple[float, float]:
    """Calcula largura/altura (em polegadas) que preservam a proporção da imagem dentro de uma caixa máxima."""
    with PILImage.open(BytesIO(png_bytes)) as image:
        pixel_width, pixel_height = image.size
    aspect = pixel_width / pixel_height
    width_in, height_in = max_width_in, max_width_in / aspect
    if height_in > max_height_in:
        height_in, width_in = max_height_in, max_height_in * aspect
    return width_in, height_in


def _build_analytics_pptx(charts: dict[str, alt.Chart]) -> bytes:
    """Monta uma apresentação PPTX com um slide por gráfico de analytics, mantendo a proporção original."""
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]
    slide_width_in = presentation.slide_width / Inches(1)
    slide_height_in = presentation.slide_height / Inches(1)
    content_top_in = 1.2
    content_bottom_margin_in = 0.4
    content_side_margin_in = 0.5
    max_width_in = slide_width_in - 2 * content_side_margin_in
    max_height_in = slide_height_in - content_top_in - content_bottom_margin_in
    for title, chart in charts.items():
        slide = presentation.slides.add_slide(blank_layout)
        title_box = slide.shapes.add_textbox(
            Inches(0.4), Inches(0.2), presentation.slide_width - Inches(0.8), Inches(0.7)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_frame.paragraphs[0].font.size = Pt(24)
        title_frame.paragraphs[0].font.bold = True
        png_bytes = _chart_to_png(chart)
        width_in, height_in = _fit_box(png_bytes, max_width_in, max_height_in)
        left_in = (slide_width_in - width_in) / 2
        top_in = content_top_in + (max_height_in - height_in) / 2
        slide.shapes.add_picture(
            BytesIO(png_bytes), Inches(left_in), Inches(top_in), width=Inches(width_in), height=Inches(height_in)
        )
    buffer = BytesIO()
    presentation.save(buffer)
    return buffer.getvalue()


def _build_analytics_pdf(charts: dict[str, alt.Chart]) -> bytes:
    """Monta um PDF com um gráfico de analytics por página, mantendo a proporção original."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    max_width_in = (letter[0] / inch) - 2
    max_height_in = 5.0
    elements = []
    for title, chart in charts.items():
        elements.append(Paragraph(title, styles["Heading1"]))
        elements.append(Spacer(1, 0.2 * inch))
        png_bytes = _chart_to_png(chart)
        width_in, height_in = _fit_box(png_bytes, max_width_in, max_height_in)
        image = RLImage(BytesIO(png_bytes), width=width_in * inch, height=height_in * inch)
        image.hAlign = "CENTER"
        elements.append(image)
        elements.append(Spacer(1, 0.4 * inch))
    doc.build(elements)
    return buffer.getvalue()
# ---------------------------------------------------------------------------
# Variáveis
# ---------------------------------------------------------------------------
status_labels = {
    "RECRUITING": "Recrutando",
    "COMPLETED": "Concluído",
    "NOT_YET_RECRUITING": "Ainda não recrutando",
    "ACTIVE_NOT_RECRUITING": "Ativo (sem recrutamento)",
    "TERMINATED": "Encerrado precocemente",
    "SUSPENDED": "Suspenso",
    "WITHDRAWN": "Retirado",
}

tipo_financiamento = {
    "INDUSTRY": "Indústria",
    "NIH": "U.S. National Institutes of Health",
    "FED": "U.S. Federal Government",
    "OTHER_GOV": "Other U.S. Federal agencies",
    "INDIV": "Indivídual",
    "NETWORK": "Rede",
    "AMBIG": "Ambíguo",
    "OTHER": "Outros",
    "UNKNOWN": "Desconhecido"
}

intervention_type_labels = {
    "BEHAVIORAL": "Comportamental",
    "BIOLOGICAL": "Biológica",
    "COMBINATION_PRODUCT": "Produto combinado",
    "DEVICE": "Dispositivos",
    "DIAGNOSTIC_TEST": "Teste diagnóstico",
    "DIETARY_SUPPLEMENT": "Suplemento dietético",
    "DRUG": "Medicamento",
    "GENETIC": "Genética",
    "PROCEDURE": "Procedimento",
    "RADIATION": "Radiação",
    "OTHER": "Outros",
}

study_type_labels = {
    "INTERVENTIONAL": "Intervencional",
    "OBSERVATIONAL": "Observacional",
    "EXPANDED_ACCESS": "Acesso expandido",
}

phase_labels = {
    "NA": "Não aplicável",
    "EARLY_PHASE1": "Fase inicial 1",
    "PHASE1": "Fase 1",
    "PHASE2": "Fase 2",
    "PHASE3": "Fase 3",
    "PHASE4": "Fase 4",
}

gender_labels = {
    "ALL": "Todos",
    "MALE": "Masculino",
    "FEMALE": "Feminino",
}



# ---------------------------------------------------------------------------
# Formulário de busca
# ---------------------------------------------------------------------------

with st.form("search_form"):
    st.subheader("1. Parâmetros de busca")
    col1, col2 = st.columns(2)
    with col1:
        query_cond = st.text_input("Condição / doença", placeholder="ex: lung cancer")
        cond_list_file = st.file_uploader(
            "Lista de condições/doenças",
            type=["txt", "xlsx", "xls"],
            help="TXT: uma condição por linha. XLSX/XLS: coluna obrigatória 'condicao'.",
        )
        query_intr = st.text_input("Intervenção / tratamento")
        intr_list_file = st.file_uploader(
            "Lista de intervenções/tratamentos",
            type=["txt", "xlsx", "xls"],
            help="TXT: uma intervenção por linha. XLSX/XLS: coluna obrigatória 'intervencao'.",
        )
        query_spons = st.text_input("Patrocinador/Colaborador")
        sponsor_list_file = st.file_uploader(
            "Lista de patrocinadores e/ou colaboradores",
            type=["txt", "xlsx", "xls"],
            help="TXT: uma empresa por linha. XLSX/XLS: coluna obrigatória 'empresa'.",
        )
        query_lead = st.text_input("Patrocinador")
        lead_list_file = st.file_uploader(
            "Lista de patrocinadores (lead)",
            type=["txt", "xlsx", "xls"],
            help="TXT: um patrocinador por linha. XLSX/XLS: coluna obrigatória 'patrocinador'.",
        )
    with col2:
        query_titles = st.text_input("Título / acrônimo")
        query_term = st.text_input(
            "Termos gerais",
            placeholder="ex: AREA[LastUpdatePostDate]RANGE[2023-01-01,MAX]",
        )
        term_list_file = st.file_uploader(
            "Lista de termos gerais",
            type=["txt", "xlsx", "xls"],
            help="TXT: um termo por linha. XLSX/XLS: coluna obrigatória 'termo'.",
        )
        query_id = st.text_input("IDs de estudo", placeholder="ex: NCT04852770")

    st.subheader("2. Filtros")
    col3, col4 = st.columns(2)
    with col3:
        filter_status = st.multiselect(
            "Status",
            options=list(status_labels.keys()),
            format_func=lambda x: status_labels.get(x, x),
            help="Ex: selecione RECRUITING e NOT_YET_RECRUITING para estudos que já recrutam ou irão recrutar.",
        )
        filter_sponsor_type = st.multiselect(
            "Tipo de patrocinador",
            options=list(tipo_financiamento.keys()),
            format_func=lambda x: tipo_financiamento.get(x, x),
        )
        with st.expander("ℹ️ Legenda dos tipos de financiamento"):
            st.markdown("""
            | Código | Nome Oficial | Descrição |
            | :--- | :--- | :--- |
            | **INDUSTRY** | Industry | Indústria farmacêutica, biotec ou dispositivos |
            | **NIH** | U.S. National Institutes of Health | Institutos Nacionais de Saúde dos EUA |
            | **FED** | U.S. Federal Gov | Agências federais dos EUA (FDA, CDC, DoD) |
            | **OTHER_GOV** | Other U.S. Federal agencies | Governos estaduais, locais ou estrangeiros |
            | **INDIV** | Individual | Pessoa física / patrocinador independente |
            | **NETWORK** | Network | Consórcios e redes colaborativas |
            | **AMBIG** | Ambiguous | Classificação ambígua |
            | **OTHER** | Other | Universidades, hospitais acadêmicos, ONGs |
            | **UNKNOWN** | Unknown | Desconhecido ou não informado |
            """)
        filter_intervention_type = st.multiselect(
            "Tipo de intervenção",
            options=list(intervention_type_labels.keys()),
            format_func=lambda x: intervention_type_labels.get(x, x),
        )
        filter_study_type = st.multiselect(
            "Tipo de estudo",
            options=list(study_type_labels.keys()),
            format_func=lambda x: study_type_labels.get(x, x),
        )
        filter_ids_raw = st.text_input("NCT IDs específicos", placeholder="vírgula p/ múltiplos")
    with col4:
        filter_phase = st.multiselect(
            "Fase",
            options=list(phase_labels.keys()),
            format_func=lambda x: phase_labels.get(x, x),
        )
        location_country_operator = st.selectbox(
            "País da localidade",
            options=["Contém", "Não contém"],
            help="Ex: selecione 'Não contém' e informe Brasil para excluir estudos com locais no Brasil.",
        )
        location_country = st.text_input(
            "País da localidade",
            placeholder="ex: Brazil",
        )
        filter_sex = st.multiselect(
            "Sexo elegível (Sex)",
            options=list(gender_labels.keys()),
            format_func=lambda x: gender_labels.get(x, x),
        )
        filter_advanced = st.text_input("Filtro avançado Essie adicional", placeholder="ex: AREA[StartDate]2022")
        filter_geo = st.text_input("Filtro geográfico", placeholder="ex: distance(-23.55,-46.63,50km)")
        st.markdown("**Enrollment (meta de recrutamento)**")
        enrollment_min_col, enrollment_max_col = st.columns(2)
        with enrollment_min_col:
            filter_enrollment_min = st.number_input("Mínimo", min_value=0, value=0, step=10)
        with enrollment_max_col:
            filter_enrollment_max = st.number_input("Máximo", min_value=0, value=0, step=10, help="0 = sem limite máximo")


    st.subheader("3. Campos e ordenação")
    custom_fields_raw = st.text_input(
        "Campos customizados (vírgula p/ múltiplos, vazio = padrão)",
        placeholder=", ".join(DEFAULT_FIELDS),
    )
    sort_raw = st.text_input("Ordenação", placeholder="ex: LastUpdatePostDate | EnrollmentCount:desc")

    st.subheader("4. Saída")
    col5, col6 = st.columns(2)
    with col5:
        max_studies_raw = st.text_input("Máximo de estudos a extrair (vazio = todos)")
    with col6:
        output_format = st.selectbox("Formato de saída", options=["csv", "json", "xlsx"], index=0)

    submitted = st.form_submit_button("🔎 Buscar")

# ---------------------------------------------------------------------------
# Executar ETL
# ---------------------------------------------------------------------------

if submitted:
    filter_ids_input = [v.strip() for v in filter_ids_raw.split(",") if v.strip()] or None
    custom_fields = [v.strip() for v in custom_fields_raw.split(",") if v.strip()] or None
    fields = custom_fields if custom_fields else DEFAULT_FIELDS
    fields = list(dict.fromkeys([
        *fields,
        "InterventionName",
        "InterventionType",
        "InterventionDescription",
        "LocationCountry",
        "CentralContactName",
        "CentralContactRole",
        "CentralContactPhone",
        "CentralContactEMail",
        "CollaboratorName",
        "LocationGeoPoint",
    ]))
    sort = [v.strip() for v in sort_raw.split(",") if v.strip()] or None
    max_studies = int(max_studies_raw) if max_studies_raw.strip().isdigit() else None

    advanced_clauses = [
        c for c in (
            _essie_clause("LeadSponsorClass", filter_sponsor_type),
            _essie_clause("InterventionType", filter_intervention_type),
            _essie_clause("StudyType", filter_study_type),
            _essie_clause("Phase", filter_phase),
            _essie_clause("Sex", filter_sex),
            _enrollment_clause(filter_enrollment_min, filter_enrollment_max),
            (
                f"{'NOT ' if location_country_operator == 'Não contém' else ''}"
                f"AREA[LocationCountry]{location_country}"
                if location_country.strip()
                else None
            ),
            filter_advanced or None,
        )
        if c
    ]
    combined_advanced = " AND ".join(advanced_clauses) or None

    def _collect_values(manual_value: str, uploaded_file, column: str) -> list[str]:
        """Combina o valor digitado manualmente com os valores de um arquivo de lista, sem duplicatas."""
        values = [manual_value] if manual_value.strip() else []
        if uploaded_file is not None:
            values.extend(load_values_from_file(uploaded_file, filename=uploaded_file.name, column=column))
        seen: set[str] = set()
        unique_values: list[str] = []
        for value in values:
            key = value.strip().casefold()
            if value.strip() and key not in seen:
                unique_values.append(value.strip())
                seen.add(key)
        return unique_values

    with st.spinner("Extraindo dados do ClinicalTrials.gov..."):
        try:
            search_kwargs = {
                "query_cond": query_cond or None,
                "query_intr": query_intr or None,
                "query_lead": query_lead or None,
                "query_titles": query_titles or None,
                "query_term": query_term or None,
                "query_id": query_id or None,
                "filter_overall_status": filter_status or None,
                "filter_ids": filter_ids_input,
                "filter_advanced": combined_advanced,
                "filter_geo": filter_geo or None,
                "sort": sort,
            }
            sponsors = [query_spons] if query_spons.strip() else []
            if sponsor_list_file is not None:
                sponsors.extend(
                    load_sponsors_from_file(
                        sponsor_list_file,
                        filename=sponsor_list_file.name,
                    )
                )

            # Lista de eixos de lote suportados: cada um combina campo de busca,
            # arquivo enviado e o rótulo usado no resumo por valor consultado.
            list_axes = [
                ("query_cond", cond_list_file, "condicao", "condicao_consultada",
                 _collect_values(query_cond, cond_list_file, "condicao")),
                ("query_intr", intr_list_file, "intervencao", "intervencao_consultada",
                 _collect_values(query_intr, intr_list_file, "intervencao")),
                ("query_term", term_list_file, "termo", "termo_consultado",
                 _collect_values(query_term, term_list_file, "termo")),
                ("query_lead", lead_list_file, "patrocinador", "patrocinador_consultado",
                 _collect_values(query_lead, lead_list_file, "patrocinador")),
            ]
            file_driven_axes = [axis for axis in list_axes if axis[1] is not None]

            if sponsors:
                records, batch_summary = run_sponsor_batch(
                    sponsors,
                    max_studies_per_sponsor=max_studies,
                    fields=fields,
                    **search_kwargs,
                )
                st.session_state["sponsor_summary"] = batch_summary
            elif file_driven_axes:
                if len(file_driven_axes) > 1:
                    st.warning(
                        "Apenas uma lista é processada por busca; considerando somente a primeira lista informada."
                    )
                query_field, _, _, summary_label, values = file_driven_axes[0]
                batch_kwargs = {key: value for key, value in search_kwargs.items() if key != query_field}
                records, batch_summary = run_multi_value_batch(
                    values,
                    query_field=query_field,
                    max_studies_per_value=max_studies,
                    fields=fields,
                    summary_label=summary_label,
                    **batch_kwargs,
                )
                st.session_state["sponsor_summary"] = batch_summary
            else:
                records = run_etl(
                    output_path=f"results.{output_format}",  # não utilizado diretamente; download é feito via UI
                    output_format="dataframe",
                    max_studies=max_studies,
                    fields=fields,
                    **search_kwargs,
                )
                st.session_state.pop("sponsor_summary", None)
            records = _dedupe_by_consulted_values(records)
            st.session_state["df"] = records
            st.session_state["output_format"] = output_format
        except Exception as e:
            st.error(f"Erro ao extrair dados: {e}")

# ---------------------------------------------------------------------------
# Resultado e download
# ---------------------------------------------------------------------------

df = st.session_state.get("df")
if df is not None:
    st.subheader("5. Resultado")
    st.write(f"**{len(df)}** estudos encontrados.")
    sponsor_summary = st.session_state.get("sponsor_summary")
    if sponsor_summary is not None:
        with st.expander("Resumo por valor consultado"):
            st.dataframe(sponsor_summary, width="stretch", hide_index=True)
    display_df = df.drop(columns=["_site_locations"], errors="ignore").copy()
    priority_columns = [
        column for column in ("nct_id", "brief_title", "acronym", "countries")
        if column in display_df.columns
    ]
    ordered_columns = priority_columns + [
        column for column in display_df.columns if column not in priority_columns
    ]
    display_df = display_df[ordered_columns]
    if "nct_id" in display_df.columns:
        display_df["nct_id"] = display_df["nct_id"].map(
            lambda nct_id: f"https://clinicaltrials.gov/study/{nct_id}"
            if pd.notna(nct_id) and str(nct_id).strip()
            else ""
        )
        st.dataframe(
            display_df,
            column_config={
                "nct_id": st.column_config.LinkColumn(
                    "NCT ID",
                    display_text=r"https://clinicaltrials.gov/study/(.*)",
                ),
                "brief_title": st.column_config.TextColumn("Título"),
                "acronym": st.column_config.TextColumn("Acrônimo"),
                "countries": st.column_config.TextColumn("Países"),
                "collaborator_names": st.column_config.TextColumn("Empresas colaboradoras"),
                "central_contact_names": st.column_config.TextColumn("Contato central - nome"),
                "central_contact_roles": st.column_config.TextColumn("Contato central - função"),
                "central_contact_phones": st.column_config.TextColumn("Contato central - telefone"),
                "central_contact_emails": st.column_config.TextColumn("Contato central - e-mail"),
            },
            width="stretch",
        )
    else:
        st.dataframe(display_df, width="stretch")

    output_format = st.session_state.get("output_format", "csv")
    export_df = df.drop(columns=["_site_locations"], errors="ignore")
    export_priority_columns = [
        column for column in ("nct_id", "brief_title", "acronym", "countries")
        if column in export_df.columns
    ]
    export_df = export_df[
        export_priority_columns
        + [column for column in export_df.columns if column not in export_priority_columns]
    ]
    if output_format == "csv":
        data = export_df.to_csv(index=False).encode("utf-8")
        mime = "text/csv"
        file_name = "results.csv"
    elif output_format == "json":
        data = json.dumps(export_df.to_dict(orient="records"), ensure_ascii=False, indent=2).encode("utf-8")
        mime = "application/json"
        file_name = "results.json"
    elif output_format == "xlsx":
        from io import BytesIO

        output = BytesIO()
        export_df.to_excel(output, index=False, engine="openpyxl")
        data = output.getvalue()
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        file_name = "results.xlsx"

    st.download_button(
        label=f"⬇️ Baixar {output_format.upper()}",
        data=data,
        file_name=file_name,
        mime=mime,
    )

    # -----------------------------------------------------------------
    # Analytics
    # -----------------------------------------------------------------
    st.divider()
    st.subheader("6. Analytics")

    if "geo_filter" not in st.session_state:
        st.session_state["geo_filter"] = "Mundo"

    geo_col1, geo_col2 = st.columns(2)
    with geo_col1:
        if st.button("🇧🇷 Brasil", width="stretch"):
            st.session_state["geo_filter"] = "Brasil"
    with geo_col2:
        if st.button("🌎 Mundo", width="stretch"):
            st.session_state["geo_filter"] = "Mundo"

    if st.session_state["geo_filter"] == "Brasil" and "countries" in df.columns:
        df = df[df["countries"].fillna("").astype(str).str.contains("Brazil", case=False)].copy()

    st.caption(f"Filtro geográfico atual: **{st.session_state['geo_filter']}** ({len(df)} estudos)")

    analytics_charts: dict[str, alt.Chart] = {}

    def _split_column(column: str) -> pd.DataFrame:
        """Transforma valores separados por ponto e vírgula em linhas contáveis."""
        if column not in df.columns:
            return pd.DataFrame(columns=[column])
        values = df[column].fillna("").astype(str).str.split(";").explode().str.strip()
        return values[values.ne("")].to_frame(name=column)

    def _show_unavailable(message: str) -> None:
        st.info(message)

    def _site_location_frame() -> pd.DataFrame:
        """Converte os locais geocodificados armazenados por estudo em pontos do mapa."""
        if "_site_locations" not in df.columns:
            return pd.DataFrame(columns=["latitude", "longitude"])

        map_points = []
        for nct_id, locations_json in df[["nct_id", "_site_locations"]].itertuples(index=False):
            try:
                locations = json.loads(locations_json or "[]")
            except (TypeError, json.JSONDecodeError):
                continue
            for location in locations:
                map_points.append({"nct_id": nct_id, **location})
        return pd.DataFrame(map_points)

    # 1. Cronograma: volume por ano e duração esperada do estudo.
    st.markdown("**1. Análises temporais e de cronograma**")
    start_dates = pd.to_datetime(
        df["start_date"] if "start_date" in df.columns else pd.Series(index=df.index),
        errors="coerce",
    )
    completion_dates = pd.to_datetime(
        df["primary_completion_date"]
        if "primary_completion_date" in df.columns
        else pd.Series(index=df.index),
        errors="coerce",
    )
    duration_days = (completion_dates - start_dates).dt.days
    duration = duration_days[duration_days.ge(0)]
    temporal_col, duration_col = st.columns(2)
    with temporal_col:
        if start_dates.notna().any():
            starts_by_year = (
                start_dates.dropna().dt.year.value_counts().sort_index()
                .rename_axis("ano").reset_index(name="estudos")
            )
            st.markdown("**Volume de início por ano**")
            starts_chart = (
                alt.Chart(starts_by_year)
                .mark_line(point=True)
                .encode(x=alt.X("ano:O", title="Ano"), y=alt.Y("estudos:Q", title="Estudos"))
            )
            st.altair_chart(starts_chart, use_container_width=True)
            analytics_charts["Volume de início por ano"] = starts_chart
        else:
            _show_unavailable("Não há datas de início válidas para esta análise.")
    with duration_col:
        if duration.size:
            st.metric("Duração média esperada", f"{duration.mean() / 30.44:.1f} meses", border=True)
            st.metric("Estudos com datas comparáveis", f"{duration.size}", border=True)
            duration_df = pd.DataFrame({"duração (meses)": duration / 30.44})
            st.bar_chart(duration_df, x=None, y="duração (meses)", x_label="Estudo", y_label="Meses")
        else:
            _show_unavailable("Não há pares válidos de início e conclusão primária.")

    # 2. Enrollment: escala geral e relação com fase clínica.
    st.markdown("**2. Análises de volumetria e escopo**")
    enrollment = pd.to_numeric(
        df["enrollment_count"]
        if "enrollment_count" in df.columns
        else pd.Series(index=df.index),
        errors="coerce",
    ).dropna()
    enrollment_col, phase_enrollment_col = st.columns(2)
    with enrollment_col:
        if enrollment.size:
            metrics = st.container(horizontal=True)
            with metrics:
                st.metric("Média", f"{enrollment.mean():,.0f}", border=True)
                st.metric("Mediana", f"{enrollment.median():,.0f}", border=True)
                st.metric("Máximo", f"{enrollment.max():,.0f}", border=True)
            st.caption(f"Base válida: {enrollment.size} de {len(df)} estudos.")
        else:
            _show_unavailable("Não há valores numéricos de enrollment.")
    with phase_enrollment_col:
        if "phases" in df.columns and enrollment.size:
            phase_data = df[["phases", "enrollment_count"]].copy()
            phase_data["enrollment_count"] = pd.to_numeric(phase_data["enrollment_count"], errors="coerce")
            phase_data["phases"] = phase_data["phases"].fillna("").astype(str).str.split(";")
            phase_data = phase_data.explode("phases")
            phase_data["phases"] = phase_data["phases"].str.strip()
            phase_data = phase_data[phase_data["phases"].ne("")]
            phase_data["enrollment_count"] = pd.to_numeric(phase_data["enrollment_count"], errors="coerce")
            phase_summary = (
                phase_data.dropna(subset=["enrollment_count"])
                .groupby("phases", as_index=False)["enrollment_count"].mean()
                .rename(columns={"enrollment_count": "média de pacientes"})
            )
            st.markdown("**Enrollment médio por fase**")
            phase_chart = (
                alt.Chart(phase_summary)
                .mark_bar()
                .encode(x=alt.X("phases:N", title="Fase"), y=alt.Y("média de pacientes:Q", title="Pacientes"))
            )
            st.altair_chart(phase_chart, use_container_width=True)
            analytics_charts["Enrollment médio por fase"] = phase_chart
        else:
            _show_unavailable("Não há fases e enrollment válidos para cruzamento.")

    # 3. Geografia: concentração e complexidade logística.
    st.markdown("**3. Análises geográficas e logísticas**")
    countries = _split_column("countries")
    if "countries" in df.columns:
        country_counts = countries["countries"].value_counts().head(15).rename_axis("país").reset_index(name="estudos")
        country_col, complexity_col = st.columns(2)
        with country_col:
            st.markdown("**Polos de pesquisa: top 15 países**")
            country_chart = (
                alt.Chart(country_counts)
                .mark_bar()
                .encode(
                    x=alt.X("estudos:Q", title="Estudos"),
                    y=alt.Y("país:N", sort="-x", title="País"),
                )
            )
            st.altair_chart(country_chart, use_container_width=True)
            analytics_charts["Polos de pesquisa: top 15 países"] = country_chart
        with complexity_col:
            country_count_per_study = df["countries"].fillna("").astype(str).map(
                lambda value: len({country.strip() for country in value.split(";") if country.strip()})
            )
            country_count_per_study = country_count_per_study[country_count_per_study.gt(0)]
            if country_count_per_study.size:
                st.metric("Média de países por estudo", f"{country_count_per_study.mean():.1f}", border=True)
                st.metric("Máximo de países em um estudo", f"{country_count_per_study.max()}", border=True)
                st.caption(f"Base válida: {country_count_per_study.size} estudos.")
            else:
                _show_unavailable("Não há países válidos para medir a complexidade logística.")
    else:
        _show_unavailable("A coluna countries não está disponível.")

    map_locations = _site_location_frame()
    if not map_locations.empty:
        st.markdown("**Mapa dos centros de pesquisa**")
        st.caption(f"{len(map_locations)} centros com latitude e longitude disponíveis.")
        st.map(map_locations, latitude="latitude", longitude="longitude", height=520)
    else:
        _show_unavailable("Não há coordenadas disponíveis para exibir os centros no mapa.")

    # 4. Patrocinadores: participação e perfil operacional.
    st.markdown("**4. Análises de mercado e patrocinadores**")
    sponsor_columns = ["lead_sponsor_name"]
    if "enrollment_count" in df.columns:
        sponsor_columns.append("enrollment_count")
    sponsor_data = df[sponsor_columns].copy() if "lead_sponsor_name" in df.columns else pd.DataFrame()
    if not sponsor_data.empty:
        sponsor_data["lead_sponsor_name"] = sponsor_data["lead_sponsor_name"].fillna("").astype(str).str.strip()
        if "enrollment_count" in sponsor_data.columns:
            sponsor_data["enrollment_count"] = pd.to_numeric(sponsor_data["enrollment_count"], errors="coerce")
        else:
            sponsor_data["enrollment_count"] = float("nan")
        sponsor_data = sponsor_data[sponsor_data["lead_sponsor_name"].ne("")]
        sponsor_summary = (
            sponsor_data.groupby("lead_sponsor_name", as_index=False)
            .agg(estudos=("lead_sponsor_name", "size"), enrollment_médio=("enrollment_count", "mean"))
            .sort_values("estudos", ascending=False).head(15)
        )
        sponsor_chart = (
            alt.Chart(sponsor_summary)
            .mark_bar()
            .encode(
                x=alt.X("estudos:Q", title="Estudos"),
                y=alt.Y("lead_sponsor_name:N", sort="-x", title="Patrocinador"),
            )
        )
        st.altair_chart(sponsor_chart, use_container_width=True)
        st.dataframe(sponsor_summary, hide_index=True, width="stretch")
        analytics_charts["Top 15 patrocinadores"] = sponsor_chart
    else:
        _show_unavailable("Não há patrocinadores válidos para esta análise.")

    # 5. Perfil clínico e científico: condições e tipos de intervenção.
    st.markdown("**5. Análises clínicas e científicas**")
    condition_data = _split_column("conditions")
    intervention_data = _split_column("intervention_types")
    condition_col, intervention_col = st.columns(2)
    with condition_col:
        if not condition_data.empty:
            condition_counts = condition_data["conditions"].value_counts().head(15).rename_axis("condição").reset_index(name="estudos")
            st.markdown("**Condições mais frequentes**")
            condition_chart = (
                alt.Chart(condition_counts)
                .mark_bar()
                .encode(
                    x=alt.X("estudos:Q", title="Estudos"),
                    y=alt.Y("condição:N", sort="-x", title="Condição"),
                )
            )
            st.altair_chart(condition_chart, use_container_width=True)
            analytics_charts["Condições mais frequentes"] = condition_chart
        else:
            _show_unavailable("Não há condições válidas para análise.")
    with intervention_col:
        if not intervention_data.empty:
            intervention_counts = intervention_data["intervention_types"].value_counts().rename_axis("tipo").reset_index(name="ocorrências")
            st.markdown("**Tipos de intervenção**")
            intervention_chart = (
                alt.Chart(intervention_counts)
                .mark_bar()
                .encode(x=alt.X("tipo:N", title="Tipo"), y=alt.Y("ocorrências:Q", title="Ocorrências"))
            )
            st.altair_chart(intervention_chart, use_container_width=True)
            analytics_charts["Tipos de intervenção"] = intervention_chart
            intervention_profile = df["intervention_types"].fillna("").astype(str).map(
                lambda value: [item.strip() for item in value.split(";") if item.strip()]
            )
            intervention_profile = intervention_profile[intervention_profile.map(bool)]
            combined_share = intervention_profile.map(lambda items: len(set(items)) > 1).mean() * 100
            drug_share = intervention_profile.map(lambda items: all(item == "DRUG" for item in items)).mean() * 100
            metrics = st.container(horizontal=True)
            with metrics:
                st.metric("Estudos com intervenção combinada", f"{combined_share:.1f}%", border=True)
                st.metric("Estudos somente com medicamentos", f"{drug_share:.1f}%", border=True)
        else:
            _show_unavailable("Não há tipos de intervenção válidos para análise.")

    # -----------------------------------------------------------------
    # Exportação dos gráficos de analytics
    # -----------------------------------------------------------------
    st.divider()
    st.subheader("7. Exportar analytics")
    if not analytics_charts:
        st.info("Nenhum gráfico disponível para exportação com os dados atuais.")
    else:
        st.caption(f"{len(analytics_charts)} gráfico(s) serão incluídos na exportação.")
        export_col1, export_col2 = st.columns(2)
        with export_col1:
            if st.button("📊 Gerar PPTX", width="stretch"):
                st.session_state["analytics_pptx"] = _build_analytics_pptx(analytics_charts)
            if "analytics_pptx" in st.session_state:
                st.download_button(
                    label="⬇️ Baixar PPTX",
                    data=st.session_state["analytics_pptx"],
                    file_name="analytics_clinicaltrials.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    width="stretch",
                )
        with export_col2:
            if st.button("📄 Gerar PDF", width="stretch"):
                st.session_state["analytics_pdf"] = _build_analytics_pdf(analytics_charts)
            if "analytics_pdf" in st.session_state:
                st.download_button(
                    label="⬇️ Baixar PDF",
                    data=st.session_state["analytics_pdf"],
                    file_name="analytics_clinicaltrials.pdf",
                    mime="application/pdf",
                    width="stretch",
                )

