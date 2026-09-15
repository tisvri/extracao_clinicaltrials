# Bibliotecas
from __future__ import annotations

import io
import json
import openpyxl

import altair as alt
import pandas as pd
import streamlit as st

from client import (
    ClinicalTrialsClient,
    DEFAULT_FIELDS,
    run_etl,
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
        query_intr = st.text_input("Intervenção / tratamento")
        query_spons = st.text_input("Patrocinador")
        query_lead = st.text_input("Lead sponsor name")
    with col2:
        query_titles = st.text_input("Título / acrônimo")
        query_term = st.text_input(
            "Termos gerais",
            placeholder="ex: AREA[LastUpdatePostDate]RANGE[2023-01-01,MAX]",
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

    with st.spinner("Extraindo dados do ClinicalTrials.gov..."):
        try:
            records = run_etl(
                output_path=f"results.{output_format}",  # não utilizado diretamente; download é feito via UI
                output_format="dataframe",
                max_studies=max_studies,
                fields=fields,
                query_cond=query_cond or None,
                query_intr=query_intr or None,
                query_spons=query_spons or None,
                query_lead=query_lead or None,
                query_titles=query_titles or None,
                query_term=query_term or None,
                query_id=query_id or None,
                filter_overall_status=filter_status or None,
                filter_ids=filter_ids_input,
                filter_advanced=combined_advanced,
                filter_geo=filter_geo or None,
                sort=sort,
            )
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
    display_df = df.copy()
    if "countries" in display_df.columns:
        priority_columns = [
            column for column in ("nct_id", "countries")
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
                "countries": st.column_config.TextColumn("Países"),
            },
            width="stretch",
        )
    else:
        st.dataframe(display_df, width="stretch")

    output_format = st.session_state.get("output_format", "csv")
    if output_format == "csv":
        data = df.to_csv(index=False).encode("utf-8")
        mime = "text/csv"
        file_name = "results.csv"
    elif output_format == "json":
        data = json.dumps(df.to_dict(orient="records"), ensure_ascii=False, indent=2).encode("utf-8")
        mime = "application/json"
        file_name = "results.json"
    elif output_format == "xlsx":
        from io import BytesIO

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
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

    def _split_column(column: str) -> pd.DataFrame:
        """Transforma valores separados por ponto e vírgula em linhas contáveis."""
        if column not in df.columns:
            return pd.DataFrame(columns=[column])
        values = df[column].fillna("").astype(str).str.split(";").explode().str.strip()
        return values[values.ne("")].to_frame(name=column)

    def _show_unavailable(message: str) -> None:
        st.info(message)

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
            st.line_chart(starts_by_year, x="ano", y="estudos", x_label="Ano", y_label="Estudos")
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
            st.bar_chart(phase_summary, x="phases", y="média de pacientes", x_label="Fase", y_label="Pacientes")
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
            st.bar_chart(country_counts, x="país", y="estudos", x_label="País", y_label="Estudos", horizontal=True)
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
        st.bar_chart(sponsor_summary, x="lead_sponsor_name", y="estudos", x_label="Patrocinador", y_label="Estudos", horizontal=True)
        st.dataframe(sponsor_summary, hide_index=True, width="stretch")
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
            st.bar_chart(condition_counts, x="condição", y="estudos", x_label="Condição", y_label="Estudos", horizontal=True)
        else:
            _show_unavailable("Não há condições válidas para análise.")
    with intervention_col:
        if not intervention_data.empty:
            intervention_counts = intervention_data["intervention_types"].value_counts().rename_axis("tipo").reset_index(name="ocorrências")
            st.markdown("**Tipos de intervenção**")
            st.bar_chart(intervention_counts, x="tipo", y="ocorrências", x_label="Tipo", y_label="Ocorrências")
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

