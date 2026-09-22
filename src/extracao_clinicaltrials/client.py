from __future__ import annotations
import logging
import requests
import pandas as pd
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from pathlib import Path
from requests.exceptions import RequestException

# Configuração de logging para monitoramento de automação
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

@dataclass
class CTGovFilters:
    """Configuração de filtros para a API v2 do ClinicalTrials.gov."""
    query_cond: Optional[str] = None          # Ex: "Breast Cancer" ou "Diabetes"
    query_term: Optional[str] = None          # Termos gerais, intervenções ou NCTId
    filter_overallStatus: Optional[str] = None # Ex: "RECRUITING", "COMPLETED"
    page_size: int = 100                      # Máximo de 1000 na v2

    def to_params(self) -> Dict[str, Any]:
        """Converte as variáveis em parâmetros formatados para a API."""
        params = {"pageSize": self.page_size, "format": "json"}
        if self.query_cond:
            params["query.cond"] = self.query_cond
        if self.query_term:
            params["query.term"] = self.query_term
        if self.filter_overallStatus:
            params["filter.overallStatus"] = self.filter_overallStatus
        return params

class CTGovExtractor:
    """Classe responsável por orquestrar a extração e tratamento dos dados."""
    BASE_URL = "https://clinicaltrials.gov/api/v2/studies"

    def __init__(self, filters: CTGovFilters):
        self.filters = filters
        self.session = requests.Session()
        
    def _parse_study(self, study: Dict[str, Any]) -> Dict[str, Any]:
        """Achata a hierarquia JSON do estudo para o DataFrame."""
        protocol = study.get("protocolSection", {})
        ident = protocol.get("identificationModule", {})
        status = protocol.get("statusModule", {})
        conds = protocol.get("conditionsModule", {}).get("conditions", [])
        
        return {
            "nct_id": ident.get("nctId"),
            "title": ident.get("briefTitle"),
            "status": status.get("overallStatus"),
            "start_date": status.get("startDateStruct", {}).get("date"),
            "conditions": " | ".join(conds) if conds else None
        }
        """
        client.py — ClinicalTrials.gov API v2 | Camada ETL
        =====================================================
        Responsabilidades:
        - Extract  : requisições HTTP à API (com paginação automática)
        - Transform : normalização / achatamento dos dados JSON
        - Load      : exportação para CSV, JSON ou DataFrame (pandas)

        Base URL: https://clinicaltrials.gov/api/v2
        Autenticação: não necessária
        """



import csv
import json
import logging
import time
from collections.abc import Iterable
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO, Generator, Optional, TextIO

import requests

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

BASE_URL = "https://clinicaltrials.gov/api/v2"
DEFAULT_PAGE_SIZE = 100          # máximo suportado pela API: 1 000
DEFAULT_FIELDS = [
    "NCTId",
    "BriefTitle",
    "Acronym",
    "OverallStatus",
    "Phase",
    "Condition",
    "InterventionName",
    "InterventionType",
    "InterventionDescription",
    "LeadSponsorName",
    "StartDate",
    "PrimaryCompletionDate",
    "CompletionDate",
    "StudyFirstPostDate",
    "EnrollmentCount",
    "StudyType",
    "HasResults",
    "LocationCountry",
    "CentralContactName",
    "CentralContactRole",
    "CentralContactPhone",
    "CentralContactEMail",
    "CollaboratorName",
    "LocationGeoPoint",
]

logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# ---------------------------------------------------------------------------
# EXTRACT
# ---------------------------------------------------------------------------

class ClinicalTrialsClient:
    """Cliente HTTP para a API do ClinicalTrials.gov (v2)."""

    def __init__(self, page_size: int = DEFAULT_PAGE_SIZE, sleep_between_pages: float = 0.3):
        self.page_size = page_size
        self.sleep_between_pages = sleep_between_pages
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})

    # ------------------------------------------------------------------
    # Primitiva: busca uma página
    # ------------------------------------------------------------------

    def _fetch_page(
        self,
        params: dict[str, Any],
        page_token: Optional[str] = None,
    ) -> dict[str, Any]:
        """Faz GET /studies e retorna o JSON bruto de uma página."""
        request_params = {**params}
        if page_token:
            request_params["pageToken"] = page_token

        response = self.session.get(f"{BASE_URL}/studies", params=request_params, timeout=30)
        try:
            response.raise_for_status()
        except RequestException as error:
            detail = response.text.strip()
            raise RequestException(f"{error} | Resposta da API: {detail}", response=response) from error
        return response.json()

    # ------------------------------------------------------------------
    # Iterador de todas as páginas (paginação automática)
    # ------------------------------------------------------------------

    def iter_studies(
        self,
        query_cond: Optional[str] = None,
        query_term: Optional[str] = None,
        query_intr: Optional[str] = None,
        query_titles: Optional[str] = None,
        query_spons: Optional[str] = None,
        query_lead: Optional[str] = None,
        query_id: Optional[str] = None,
        filter_overall_status: Optional[list[str]] = None,
        filter_ids: Optional[list[str]] = None,
        filter_advanced: Optional[str] = None,
        filter_geo: Optional[str] = None,
        fields: Optional[list[str]] = None,
        sort: Optional[list[str]] = None,
        max_studies: Optional[int] = None,
    ) -> Generator[dict[str, Any], None, None]:
        """
        Itera sobre todos os estudos que correspondem aos parâmetros.
        Faz paginação automática até esgotar os resultados ou atingir max_studies.

        Yields
        ------
        dict
            Objeto JSON bruto de cada estudo (campo 'studies' da API).
        """
        params: dict[str, Any] = {"pageSize": self.page_size, "format": "json"}

        if query_cond:                   params["query.cond"]          = query_cond
        if query_term:                   params["query.term"]          = query_term
        if query_intr:                   params["query.intr"]          = query_intr
        if query_titles:                 params["query.titles"]        = query_titles
        if query_spons:                  params["query.spons"]         = query_spons
        if query_lead:                   params["query.lead"]          = query_lead
        if query_id:                     params["query.id"]            = query_id
        if filter_overall_status:        params["filter.overallStatus"]= ",".join(filter_overall_status)
        if filter_ids:                   params["filter.ids"]          = ",".join(filter_ids)
        if filter_advanced:              params["filter.advanced"]     = filter_advanced
        if filter_geo:                   params["filter.geo"]          = filter_geo
        if fields:                       params["fields"]              = ",".join(fields)
        if sort:                         params["sort"]                = ",".join(sort)

        total_yielded = 0
        page_token: Optional[str] = None
        page_num = 1

        while True:
            logger.info("Buscando página %d ...", page_num)
            data = self._fetch_page(params, page_token)

            studies = data.get("studies", [])
            for study in studies:
                if max_studies and total_yielded >= max_studies:
                    return
                yield study
                total_yielded += 1

            page_token = data.get("nextPageToken")
            if not page_token:
                break

            page_num += 1
            time.sleep(self.sleep_between_pages)

        logger.info("Extração concluída. Total de estudos: %d", total_yielded)

    # ------------------------------------------------------------------
    # Busca de estudo único por NCT ID
    # ------------------------------------------------------------------

    def get_study(self, nct_id: str, fields: Optional[list[str]] = None) -> dict[str, Any]:
        """Retorna o JSON completo de um estudo específico pelo NCT ID."""
        params: dict[str, Any] = {"format": "json"}
        if fields:
            params["fields"] = ",".join(fields)

        response = self.session.get(f"{BASE_URL}/studies/{nct_id}", params=params, timeout=30)
        response.raise_for_status()
        return response.json()

    # ------------------------------------------------------------------
    # Versão da API / timestamp dos dados
    # ------------------------------------------------------------------

    def get_version(self) -> dict[str, str]:
        """Retorna versão da API e timestamp da última atualização dos dados."""
        response = self.session.get(f"{BASE_URL}/version", timeout=10)
        response.raise_for_status()
        return response.json()


# ---------------------------------------------------------------------------
# TRANSFORM
# ---------------------------------------------------------------------------

def flatten_study(study: dict[str, Any]) -> dict[str, Any]:
    """
    Achata o objeto JSON de um estudo em um dicionário plano,
    adequado para exportação tabular (CSV / DataFrame).
    """
    proto = study.get("protocolSection", {})
    id_mod    = proto.get("identificationModule",  {})
    status_mod= proto.get("statusModule",          {})
    desc_mod  = proto.get("descriptionModule",     {})
    design_mod= proto.get("designModule",          {})
    enroll_mod= design_mod.get("enrollmentInfo",   {})
    sponsor_mod=proto.get("sponsorCollaboratorsModule", {})
    lead      = sponsor_mod.get("leadSponsor",     {})
    arms_mod  = proto.get("armsInterventionsModule", {})
    cond_mod  = proto.get("conditionsModule",      {})
    elig_mod  = proto.get("eligibilityModule",     {})
    contacts  = proto.get("contactsLocationsModule", {})
    central_contacts = contacts.get("centralContacts", [])
    collaborators = sponsor_mod.get("collaborators", [])

    # intervenções: pega o nome da primeira, ou todas separadas por ";"
    interventions = arms_mod.get("interventions", [])
    intervention_names = ";".join(
        i.get("name", "") for i in interventions
    ) if interventions else ""
    intervention_types = ";".join(
        i.get("type", "") for i in interventions
    ) if interventions else ""
    intervention_descriptions = ";".join(
        i.get("description", "") for i in interventions
    ) if interventions else ""

    conditions = ";".join(cond_mod.get("conditions", []))

    locations = contacts.get("locations", [])
    countries = list(dict.fromkeys(
        location.get("country", "").strip()
        for location in locations
        if location.get("country", "").strip()
    ))
    site_locations = [
        {
            "facility": location.get("facility", ""),
            "city": location.get("city", ""),
            "state": location.get("state", ""),
            "country": location.get("country", ""),
            "status": location.get("status", ""),
            "latitude": location.get("geoPoint", {}).get("lat"),
            "longitude": location.get("geoPoint", {}).get("lon"),
        }
        for location in locations
        if location.get("geoPoint", {}).get("lat") is not None
        and location.get("geoPoint", {}).get("lon") is not None
    ]

    # fases
    phases = ";".join(design_mod.get("phases", []))

    flat = {
        "nct_id":                   id_mod.get("nctId", ""),
        "brief_title":              id_mod.get("briefTitle", ""),
        "official_title":           id_mod.get("officialTitle", ""),
        "acronym":                  id_mod.get("acronym", ""),
        "overall_status":           status_mod.get("overallStatus", ""),
        "start_date":               status_mod.get("startDateStruct", {}).get("date", ""),
        "primary_completion_date":  status_mod.get("primaryCompletionDateStruct", {}).get("date", ""),
        "completion_date":          status_mod.get("completionDateStruct", {}).get("date", ""),
        "first_posted_date":        status_mod.get("studyFirstPostDateStruct", {}).get("date", ""),
        "study_type":               design_mod.get("studyType", ""),
        "phases":                   phases,
        "enrollment_count":         enroll_mod.get("count", ""),
        "enrollment_type":          enroll_mod.get("type", ""),
        "conditions":               conditions,
        "countries":                "; ".join(countries),
        "intervention_names":       intervention_names,
        "intervention_types":       intervention_types,
        "intervention_descriptions": intervention_descriptions,
        "lead_sponsor_name":        lead.get("name", ""),
        "lead_sponsor_class":       lead.get("class", ""),
        "collaborator_names":       "; ".join(
            collaborator.get("name", "")
            for collaborator in collaborators
            if collaborator.get("name")
        ),
        "central_contact_names":    "; ".join(
            contact.get("name", "") for contact in central_contacts if contact.get("name")
        ),
        "central_contact_roles":    "; ".join(
            contact.get("role", "") for contact in central_contacts if contact.get("role")
        ),
        "central_contact_phones":   "; ".join(
            contact.get("phone", "") for contact in central_contacts if contact.get("phone")
        ),
        "central_contact_emails":   "; ".join(
            contact.get("email", "") for contact in central_contacts if contact.get("email")
        ),
        "minimum_age":              elig_mod.get("minimumAge", ""),
        "maximum_age":              elig_mod.get("maximumAge", ""),
        "sex":                      elig_mod.get("sex", ""),
        "healthy_volunteers":       elig_mod.get("healthyVolunteers", ""),
        "brief_summary":            desc_mod.get("briefSummary", ""),
        "has_results":              study.get("hasResults", False),
        "_site_locations":          json.dumps(site_locations, ensure_ascii=False),
    }
    return flat


def transform_studies(studies: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Aplica flatten_study a uma lista de estudos."""
    return [flatten_study(s) for s in studies]


# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

def load_to_csv(records: list[dict[str, Any]], output_path: str | Path) -> None:
    """Salva lista de registros achatados em um arquivo CSV."""
    if not records:
        logger.warning("Nenhum registro para salvar.")
        return

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = list(records[0].keys())
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)

    logger.info("CSV salvo em: %s  (%d linhas)", output_path, len(records))


def load_to_json(records: list[dict[str, Any]], output_path: str | Path) -> None:
    """Salva lista de registros em um arquivo JSON."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    logger.info("JSON salvo em: %s  (%d registros)", output_path, len(records))


def load_to_dataframe(records: list[dict[str, Any]]):
    """Retorna um pandas DataFrame a partir dos registros achatados."""
    try:
        import pandas as pd
        df = pd.DataFrame(records)
        logger.info("DataFrame criado: %d linhas x %d colunas", *df.shape)
        return df
    except ImportError:
        raise ImportError(
            "pandas não está instalado. Execute: pip install pandas"
        )


def normalize_values(values: Iterable[object]) -> list[str]:
    """Remove valores vazios e duplicados, preservando a ordem da lista."""
    normalized: list[str] = []
    seen: set[str] = set()

    for value in values:
        if not isinstance(value, str):
            continue
        name = value.strip()
        key = name.casefold()
        if name and key not in seen:
            normalized.append(name)
            seen.add(key)

    return normalized


# Mantido por compatibilidade; patrocinadores usam a mesma normalização genérica.
normalize_sponsors = normalize_values


def load_values_from_file(
    source: str | Path | BinaryIO | TextIO,
    filename: Optional[str] = None,
    column: str = "valor",
) -> list[str]:
    """Carrega valores de TXT (um item por linha) ou XLSX/XLS.

    Em planilhas, a coluna informada por ``column`` é obrigatória.
    """
    source_name = filename or getattr(source, "name", None) or str(source)
    suffix = Path(source_name).suffix.lower()

    if suffix == ".txt":
        if isinstance(source, (str, Path)):
            content = Path(source).read_text(encoding="utf-8-sig")
        else:
            raw_content = source.read()
            content = raw_content.decode("utf-8-sig") if isinstance(raw_content, bytes) else raw_content
        return normalize_values(content.splitlines())

    if suffix in {".xlsx", ".xls"}:
        if isinstance(source, (str, Path)):
            dataframe = pd.read_excel(source)
        else:
            dataframe = pd.read_excel(BytesIO(source.read()))
        matching_columns = {str(name).strip().casefold(): name for name in dataframe.columns}
        requested_column = column.strip().casefold()
        if requested_column not in matching_columns:
            raise ValueError(f"A planilha deve conter a coluna '{column}'.")
        return normalize_values(dataframe[matching_columns[requested_column]].tolist())

    raise ValueError("Formato não suportado. Envie um arquivo .txt, .xlsx ou .xls.")


def load_sponsors_from_file(
    source: str | Path | BinaryIO | TextIO,
    filename: Optional[str] = None,
    column: str = "empresa",
) -> list[str]:
    """Carrega patrocinadores de TXT (uma empresa por linha) ou XLSX.

    Em planilhas, a coluna ``empresa`` e obrigatória por padrão. O nome pode ser
    alterado pelo parâmetro ``column``.
    """
    return load_values_from_file(source, filename=filename, column=column)


def run_multi_value_batch(
    values: Iterable[object],
    query_field: str,
    max_studies_per_value: Optional[int] = None,
    fields: Optional[list[str]] = None,
    summary_label: str = "valor_consultado",
    **search_kwargs,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Consulta a API uma vez por valor de uma lista, variando o parâmetro ``query_field``.

    Usado para buscas em lote (patrocinador, condição, intervenção ou termos gerais).
    Erros em um valor não interrompem os demais.
    """
    value_list = normalize_values(values)
    if not value_list:
        raise ValueError("A lista não contém valores válidos.")

    client = ClinicalTrialsClient()
    records: list[dict[str, Any]] = []
    summary: list[dict[str, Any]] = []

    for value in value_list:
        try:
            studies = client.iter_studies(
                fields=fields or DEFAULT_FIELDS,
                max_studies=max_studies_per_value,
                **{query_field: value},
                **search_kwargs,
            )
            value_records = transform_studies(list(studies))
            for record in value_records:
                record[summary_label] = value
            records.extend(value_records)
            summary.append({
                summary_label: value,
                "estudos_encontrados": len(value_records),
                "status_consulta": "concluida",
                "erro": "",
            })
        except RequestException as error:
            logger.exception("Falha ao consultar valor: %s", value)
            summary.append({
                summary_label: value,
                "estudos_encontrados": 0,
                "status_consulta": "erro",
                "erro": str(error),
            })

    return pd.DataFrame(records), pd.DataFrame(summary)


def run_sponsor_batch(
    sponsors: Iterable[object],
    max_studies_per_sponsor: Optional[int] = None,
    fields: Optional[list[str]] = None,
    **search_kwargs,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Consulta cada patrocinador e retorna estudos e resumo do processamento.

    ``max_studies_per_sponsor`` e aplicado individualmente, pois cada empresa
    representa uma consulta independente. Erros de uma empresa não interrompem
    as demais consultas.
    """
    return run_multi_value_batch(
        sponsors,
        query_field="query_spons",
        max_studies_per_value=max_studies_per_sponsor,
        fields=fields,
        summary_label="empresa_consultada",
        **search_kwargs,
    )


# ---------------------------------------------------------------------------
# Pipeline completo (Extract → Transform → Load)
# ---------------------------------------------------------------------------

def run_etl(
    output_path: str | Path,
    output_format: str = "csv",
    max_studies: Optional[int] = None,
    fields: Optional[list[str]] = None,
    **search_kwargs,
):
    """
    Executa o pipeline ETL completo.

    Parâmetros
    ----------
    output_path   : caminho do arquivo de saída
    output_format : 'csv' | 'json' | 'dataframe'
    max_studies   : limite de estudos a extrair (None = todos)
    fields        : campos a retornar da API (None = campos padrão)
    **search_kwargs : parâmetros de busca (ver ClinicalTrialsClient.iter_studies)

    Retorna
    -------
    list[dict] ou pandas.DataFrame (quando output_format='dataframe')
    """
    client = ClinicalTrialsClient()

    # -- Extract --
    raw_studies = list(
        client.iter_studies(
            fields=fields or DEFAULT_FIELDS,
            max_studies=max_studies,
            **search_kwargs,
        )
    )

    # -- Transform --
    records = transform_studies(raw_studies)

    # -- Load --
    if output_format == "csv":
        load_to_csv(records, output_path)
    elif output_format == "json":
        load_to_json(records, output_path)
    elif output_format == "dataframe":
        return load_to_dataframe(records)
    else:
        raise ValueError(f"output_format inválido: '{output_format}'. Use 'csv', 'json' ou 'dataframe'.")

    return records

    def fetch_data(self, max_pages: int = 5) -> pd.DataFrame:
        """Extrai dados com paginação e tratamento de falhas."""
        params = self.filters.to_params()
        all_studies: List[Dict[str, Any]] = []
        page_token = None
        
        for page in range(max_pages):
            if page_token:
                params["pageToken"] = page_token
                
            try:
                logging.info(f"Extraindo página {page + 1}... Parâmetros: {params}")
                response = self.session.get(self.BASE_URL, params=params, timeout=15)
                response.raise_for_status()
                data = response.json()
                
                studies = data.get("studies", [])
                if not studies:
                    logging.info("Nenhum estudo adicional encontrado.")
                    break
                    
                all_studies.extend([self._parse_study(s) for s in studies])
                
                page_token = data.get("nextPageToken")
                if not page_token:
                    logging.info("Fim da paginação.")
                    break
                    
            except RequestException as e:
                logging.error(f"Falha de conexão na página {page + 1}: {e}")
                break # Interrompe fluxo em caso de falha de rede persistente
                
        df = pd.DataFrame(all_studies)
        logging.info(f"Extração concluída. Total de registros: {len(df)}")
        return df

    def export(self, df: pd.DataFrame, filename: str, output_dir: str = "exports") -> None:
        """Exporta os dados sanitizados para arquivos planos."""
        if df.empty:
            logging.warning("DataFrame vazio. Abortando exportação.")
            return

        out_path = Path(output_dir)
        out_path.mkdir(parents=True, exist_ok=True)
        
        csv_path = out_path / f"{filename}.csv"
        excel_path = out_path / f"{filename}.xlsx"
        
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        df.to_excel(excel_path, index=False, engine='openpyxl')
        logging.info(f"Arquivos exportados para: {out_path.absolute()}")