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

    with st.container(horizontal=True):
        st.metric("Total de estudos", f"{len(df)}", border=True)
        if "has_results" in df.columns and len(df):
            pct_results = df["has_results"].astype(bool).sum() / len(df) * 100
            st.metric("Com resultados", f"{pct_results:.1f}%", border=True)
        if "enrollment_count" in df.columns:
            enrollment = pd.to_numeric(df["enrollment_count"], errors="coerce")
            avg_enrollment = f"{enrollment.mean():.0f}" if enrollment.notna().any() else "—"
            st.metric("Inscrição média", avg_enrollment, border=True)
        if "lead_sponsor_name" in df.columns:
            st.metric("Patrocinadores únicos", f"{df['lead_sponsor_name'].nunique()}", border=True)

    col_a, col_b = st.columns(2)
    with col_a:
        with st.container(border=True):
            st.markdown("**Distribuição por status**")
            if "overall_status" in df.columns and len(df):
                status_counts = df["overall_status"].value_counts().reset_index()
                status_counts.columns = ["status", "contagem"]
                st.bar_chart(status_counts, x="status", y="contagem", x_label="", y_label="", horizontal=True)
            else:
                st.info("Coluna 'overall_status' não disponível.")

    with col_b:
        with st.container(border=True):
            st.markdown("**Distribuição por tipo de estudo**")
            if "study_type" in df.columns and len(df):
                type_counts = df["study_type"].value_counts().reset_index()
                type_counts.columns = ["study_type", "contagem"]
                pie_chart = (
                    alt.Chart(type_counts)
                    .mark_arc(innerRadius=50)
                    .encode(
                        theta="contagem:Q",
                        color="study_type:N",
                        tooltip=["study_type", "contagem"],
                    )
                )
                st.altair_chart(pie_chart)
            else:
                st.info("Coluna 'study_type' não disponível.")

    col_c, col_d = st.columns(2)
    with col_c:
        with st.container(border=True):
            st.markdown("**Distribuição por fase**")
            if "phases" in df.columns and len(df):
                phases_series = df["phases"].replace("", "Não informado").fillna("Não informado")
                phase_counts = phases_series.value_counts().reset_index()
                phase_counts.columns = ["fase", "contagem"]
                st.bar_chart(phase_counts, x="fase", y="contagem", x_label="", y_label="")
            else:
                st.info("Coluna 'phases' não disponível.")

    with col_d:
        with st.container(border=True):
            st.markdown("**Estudos com resultados publicados**")
            if "has_results" in df.columns and len(df):
                results_counts = df["has_results"].astype(bool).value_counts().rename({True: "Sim", False: "Não"}).reset_index()
                results_counts.columns = ["tem_resultados", "contagem"]
                results_chart = (
                    alt.Chart(results_counts)
                    .mark_arc(innerRadius=50)
                    .encode(
                        theta="contagem:Q",
                        color="tem_resultados:N",
                        tooltip=["tem_resultados", "contagem"],
                    )
                )
                st.altair_chart(results_chart)
            else:
                st.info("Coluna 'has_results' não disponível.")

    with st.container(border=True):
        st.markdown("**Top 10 patrocinadores**")
        if "lead_sponsor_name" in df.columns and len(df):
            sponsor_counts = (
                df["lead_sponsor_name"]
                .replace("", "Não informado")
                .fillna("Não informado")
                .value_counts()
                .head(10)
                .reset_index()
            )
            sponsor_counts.columns = ["patrocinador", "contagem"]
            sponsor_chart = (
                alt.Chart(sponsor_counts)
                .mark_bar()
                .encode(
                    x=alt.X("contagem:Q", title=None, axis=alt.Axis(grid=False)),
                    y=alt.Y(
                        "patrocinador:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(grid=False, labelLimit=320, labelPadding=8),
                    ),
                    tooltip=["patrocinador", "contagem"],
                )
                .properties(height=320)
            )
            st.altair_chart(sponsor_chart)
        else:
            st.info("Coluna 'lead_sponsor_name' não disponível.")

