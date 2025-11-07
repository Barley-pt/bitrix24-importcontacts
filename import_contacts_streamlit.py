# app.py
import io
import json
import requests
import pandas as pd
import streamlit as st
from typing import Dict, Any, List

st.set_page_config(page_title="Importar Contatos Bitrix24", layout="wide")

# -------------------- Utilidades --------------------

def normalize_webhook(url: str) -> str:
    if not url:
        return url
    url = url.rstrip("/")
    # aceita webhook com ou sem /rest
    if not url.endswith("/rest"):
        url = url + "/rest"
    return url + "/"

def field_label(fid: str, fdata: Dict[str, Any]) -> str:
    title = (
        fdata.get("listLabel")
        or fdata.get("formLabel")
        or fdata.get("filterLabel")
        or fdata.get("title")
        or fid
    )
    return f"{title} ({fid})" if fid.upper().startswith("UF_CRM") else title

@st.cache_data(show_spinner=False)
def fetch_contact_fields(webhook: str) -> Dict[str, Any]:
    r = requests.get(f"{webhook}crm.contact.fields.json", timeout=30)
    r.raise_for_status()
    data = r.json()
    if "result" not in data:
        raise RuntimeError(f"Resposta inesperada: {data}")
    # filtra tipos graváveis e não-readonly, como no script original
    allowed = {'string','integer','double','boolean','enumeration','date','datetime'}
    fields = {
        k:v for k,v in data["result"].items()
        if v.get("type") in allowed and not v.get("isReadOnly", False)
    }
    return fields

def load_dataframe(upload) -> pd.DataFrame:
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    elif name.endswith(".xls") or name.endswith(".xlsx"):
        return pd.read_excel(upload, engine="openpyxl")
    else:
        raise ValueError("Arquivo deve ser .csv, .xls ou .xlsx")

def sanitize_value(v):
    # remove NaN e converte datas
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    return v

def ensure_multifield(lst: List[Dict[str, str]]) -> List[Dict[str, str]]:
    # remove vazios e duplicados mantendo ordem
    seen = set()
    cleaned = []
    for item in lst:
        val = (item.get("VALUE") or "").strip()
        if not val:
            continue
        key = (item.get("VALUE_TYPE") or "WORK", val)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append({"VALUE": val, "VALUE_TYPE": item.get("VALUE_TYPE") or "WORK"})
    return cleaned

def find_existing_contact(webhook: str, email: str|None, phone: str|None) -> str|None:
    # tenta por e-mail e por telefone, retorna o primeiro ID encontrado
    # Bitrix aceita filtros diretos por EMAIL/PHONE; se quiser precisão, usa "=EMAIL"
    def _query(flt: Dict[str, Any]) -> str|None:
        r = requests.post(
            f"{webhook}crm.contact.list.json",
            json={"filter": flt, "select": ["ID"]},
            timeout=30
        )
        data = r.json()
        res = data.get("result") or []
        return res[0]["ID"] if res else None

    if email:
        cid = _query({"EMAIL": email})
        if cid:
            return cid
    if phone:
        cid = _query({"PHONE": phone})
        if cid:
            return cid
    return None

def add_contact(webhook: str, fields: Dict[str, Any]) -> tuple[str|None, str]:
    r = requests.post(
        f"{webhook}crm.contact.add.json",
        json={"fields": fields, "params": {"REGISTER_SONET_EVENT": "N"}},
        timeout=60
    )
    data = r.json()
    if "result" in data and data["result"]:
        return str(data["result"]), "Created"
    return None, data.get("error_description") or json.dumps(data)

def build_payload(row: pd.Series, mapping: Dict[str,str]) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}
    emails: List[Dict[str,str]] = []
    phones: List[Dict[str,str]] = []

    for col, mapped in mapping.items():
        fid = mapped  # já é a chave Bitrix (ex: "NAME", "LAST_NAME", "EMAIL", "UF_CRM_xxx")
        raw = sanitize_value(row[col])
        if raw is None:
            continue

        if fid == "EMAIL":
            emails.append({"VALUE": str(raw), "VALUE_TYPE": "WORK"})
        elif fid == "PHONE":
            phones.append({"VALUE": str(raw), "VALUE_TYPE": "WORK"})
        else:
            fields[fid] = raw

    if emails:
        fields["EMAIL"] = ensure_multifield(emails)
    if phones:
        fields["PHONE"] = ensure_multifield(phones)
    return fields

def make_excel_with_ids(df_original: pd.DataFrame, id_list: List[str|None]) -> bytes:
    out_df = df_original.copy()
    out_df["BITRIX_ID"] = id_list
    with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="imported")
        writer._save()  # type: ignore
        data: io.BytesIO = writer.book._writer.fp  # type: ignore
    return data.getvalue()

# -------------------- UI --------------------

st.title("Importar Contatos para Bitrix24")

with st.sidebar:
    st.header("Configuração")
    webhook_in = st.text_input(
        "Webhook Bitrix24",
        placeholder="https://SEU_DOMINIO/rest/XX/YY/",
        help="Cole o webhook completo. Pode terminar com /rest/ ou apenas /; normalizamos automaticamente."
    )
    dup_check = st.toggle("Checar duplicados por e-mail/telefone", value=True,
                          help="Se ativo, pesquisa contato existente antes de criar.")
    uploaded = st.file_uploader("Arquivo (.csv, .xls, .xlsx)", type=["csv","xls","xlsx"])
    btn_fetch = st.button("Carregar campos")

# estado
if "fields" not in st.session_state: st.session_state.fields = None
if "df" not in st.session_state: st.session_state.df = None
if "mapping" not in st.session_state: st.session_state.mapping = {}

# normaliza webhook
webhook = normalize_webhook(webhook_in) if webhook_in else ""

# carrega campos
if btn_fetch:
    if not webhook:
        st.sidebar.error("Informe o webhook.")
    else:
        try:
            st.session_state.fields = fetch_contact_fields(webhook)
            st.sidebar.success("Campos carregados.")
        except Exception as e:
            st.sidebar.error(f"Falha ao obter campos: {e}")

# carrega dados
if uploaded is not None:
    try:
        st.session_state.df = load_dataframe(uploaded)
        st.success(f"Arquivo carregado com {len(st.session_state.df)} linhas.")
        st.dataframe(st.session_state.df.head(50), use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")

# mapeamento
if st.session_state.df is not None and st.session_state.fields is not None:
    st.subheader("Mapeamento de colunas → campos do Bitrix24")

    df_cols = list(st.session_state.df.columns)
    fields = st.session_state.fields

    # prepara opções ordenadas por rótulo
    sorted_items = sorted(fields.items(), key=lambda kv: field_label(kv[0], kv[1]).lower())
    labels = [field_label(fid, meta) for fid, meta in sorted_items]
    label_to_fid = {field_label(fid, meta): fid for fid, meta in sorted_items}

    left, right = st.columns(2)
    mapping: Dict[str,str] = {}

    for i, col in enumerate(df_cols):
        with (left if i % 2 == 0 else right):
            sel = st.selectbox(
                f"{col}",
                ["- não importar -"] + labels,
                index=0,
                key=f"map_{col}"
            )
            if sel != "- não importar -":
                mapping[col] = label_to_fid[sel]

    st.session_state.mapping = mapping

    st.markdown("---")
    st.subheader("Importação")
    go = st.button("Iniciar importação")

    if go:
        if not webhook:
            st.error("Informe o webhook.")
        elif not mapping:
            st.error("Defina pelo menos um mapeamento.")
        else:
            df = st.session_state.df
            total = len(df)
            progress = st.progress(0, text="Importando...")
            status = st.empty()
            logs = []
            ids = []

            for idx, row in df.iterrows():
                fields_payload = build_payload(row, mapping)

                # coleta email/phone simples para a busca (pega primeiro da lista, se existir)
                email = None
                phone = None
                if isinstance(fields_payload.get("EMAIL"), list) and fields_payload["EMAIL"]:
                    email = fields_payload["EMAIL"][0].get("VALUE")
                if isinstance(fields_payload.get("PHONE"), list) and fields_payload["PHONE"]:
                    phone = fields_payload["PHONE"][0].get("VALUE")

                try:
                    existing_id = None
                    if dup_check and (email or phone):
                        existing_id = find_existing_contact(webhook, email, phone)

                    if existing_id:
                        contact_id, result = existing_id, "DuplicateFound"
                    else:
                        contact_id, result = add_contact(webhook, fields_payload)

                    ok = bool(contact_id)
                    ids.append(contact_id)
                    logs.append({
                        "row": int(idx) + 1,
                        "result": result,
                        "contact_id": contact_id or "",
                        "payload": json.dumps(fields_payload, ensure_ascii=False)
                    })
                    status.write(f"[{len(ids)}/{total}] {'OK' if ok else 'Falha'} - ID: {contact_id or '-'}")
                except Exception as e:
                    ids.append(None)
                    logs.append({
                        "row": int(idx) + 1,
                        "result": f"Error: {e}",
                        "contact_id": "",
                        "payload": json.dumps(fields_payload, ensure_ascii=False)
                    })

                progress.progress(min(len(ids)/total, 1.0))

            # resultados
            ok_count = sum(1 for i in ids if i)
            fail_count = total - ok_count
            st.success(f"Concluído. Sucesso: {ok_count} • Falhas: {fail_count}")

            log_df = pd.DataFrame(logs, columns=["row","result","contact_id","payload"])
            st.dataframe(log_df, use_container_width=True)

            # download do log CSV
            csv_buf = io.StringIO()
            log_df.to_csv(csv_buf, index=False)
            st.download_button("Baixar log CSV", data=csv_buf.getvalue().encode("utf-8"),
                               file_name="bitrix_contacts_import_log.csv", mime="text/csv")

            # download do Excel com BITRIX_ID
            try:
                xlsx_bytes = make_excel_with_ids(df, ids)
                st.download_button(
                    "Baixar Excel com BITRIX_ID",
                    data=xlsx_bytes,
                    file_name="contacts_bitrix_imported.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.info(f"Não foi possível gerar Excel de saída: {e}")

# notas
with st.expander("Dicas e observações"):
    st.markdown(
        """
- Se a API acusar campo obrigatório não enviado, mapeie-o antes de importar.
- Para **EMAIL** e **PHONE**, podes mapear várias colunas diferentes - o app envia como multi-campos.
- O verificador de duplicados pesquisa por e-mail e/ou telefone usando `crm.contact.list`.
- O webhook deve ter permissão para `crm.contact.add`.
- Datas no Excel são convertidas para ISO (YYYY-MM-DD).
        """
    )
