# app.py
import io
import json
import requests
import pandas as pd
import streamlit as st
from typing import Dict, Any, List

st.set_page_config(page_title="Bitrix24 Contact Import-beta version", layout="wide")

# -------------------- Utilities --------------------

def normalize_webhook(url: str) -> str:
    if not url:
        return ""
    url = url.strip()
    # Do NOT append '/rest' here. The user must paste the full webhook:
    # e.g. https://fvl.bitrix24.com.br/rest/241/rct2v0wt7wair6ie/
    if not url.endswith("/"):
        url += "/"
    return url


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
    allowed = {'string','integer','double','boolean','enumeration','date','datetime','crm_multifield'}
    fields = {}
    for k, v in data["result"].items():
        if k in {"EMAIL", "PHONE"}:
            fields[k] = v
            continue
        if v.get("type") in allowed and not v.get("isReadOnly", False):
            fields[k] = v
    return fields

def load_dataframe(upload) -> pd.DataFrame:
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    elif name.endswith(".xls") or name.endswith(".xlsx"):
        return pd.read_excel(upload, engine="openpyxl")
    else:
        raise ValueError("File must be .csv, .xls or .xlsx")

def sanitize_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    return v

def ensure_multifield(lst: List[Dict[str, str]]) -> List[Dict[str, str]]:
    # remove empties and duplicates while preserving order
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
    # try by email and then by phone, return first found ID
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
        fid = mapped
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

st.title("Import Contacts to Bitrix24")

with st.sidebar:
    st.header("Setup")
    webhook_in = st.text_input(
        "Bitrix24 webhook",
        placeholder="https://YOUR_DOMAIN/rest/XX/YY/",
        help="Paste the full webhook. It can end with /rest/ or just /. It will be normalized."
    )
    dup_check = st.toggle("Check duplicates by email or phone", value=True,
                          help="If enabled, the app searches for an existing contact before creating one.")
    uploaded = st.file_uploader("File (.csv, .xls, .xlsx)", type=["csv","xls","xlsx"])
    btn_fetch = st.button("Load fields")

# state
if "fields" not in st.session_state: st.session_state.fields = None
if "df" not in st.session_state: st.session_state.df = None
if "mapping" not in st.session_state: st.session_state.mapping = {}

# normalize webhook
webhook = normalize_webhook(webhook_in) if webhook_in else ""

# load fields
if btn_fetch:
    if not webhook:
        st.sidebar.error("Enter the webhook.")
    else:
        try:
            st.session_state.fields = fetch_contact_fields(webhook)
            st.sidebar.success("Fields loaded.")
        except Exception as e:
            st.sidebar.error(f"Could not fetch fields: {e}")

# load data
if uploaded is not None:
    try:
        st.session_state.df = load_dataframe(uploaded)
        st.success(f"File loaded with {len(st.session_state.df)} rows.")
        st.dataframe(st.session_state.df.head(50), use_container_width=True)
    except Exception as e:
        st.error(f"Could not read file: {e}")

# mapping
if st.session_state.df is not None and st.session_state.fields is not None:
    st.subheader("Map columns → Bitrix24 fields")

    df_cols = list(st.session_state.df.columns)
    fields = st.session_state.fields

    sorted_items = sorted(fields.items(), key=lambda kv: field_label(kv[0], kv[1]).lower())
    labels = [field_label(fid, meta) for fid, meta in sorted_items]
    label_to_fid = {field_label(fid, meta): fid for fid, meta in sorted_items}

    left, right = st.columns(2)
    mapping: Dict[str,str] = {}

    for i, col in enumerate(df_cols):
        with (left if i % 2 == 0 else right):
            sel = st.selectbox(
                f"{col}",
                ["- do not import -"] + labels,
                index=0,
                key=f"map_{col}"
            )
            if sel != "- do not import -":
                mapping[col] = label_to_fid[sel]

    st.session_state.mapping = mapping

    st.markdown("---")
    st.subheader("Import")
    go = st.button("Start import")

    if go:
        if not webhook:
            st.error("Enter the webhook.")
        elif not mapping:
            st.error("Define at least one mapping.")
        else:
            df = st.session_state.df
            total = len(df)
            progress = st.progress(0, text="Importing...")
            status = st.empty()
            logs = []
            ids = []

            for idx, row in df.iterrows():
                fields_payload = build_payload(row, mapping)

                # pick first email or phone, if present, for duplicate check
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
                    status.write(f"[{len(ids)}/{total}] {'OK' if ok else 'Fail'} - ID: {contact_id or '-'}")
                except Exception as e:
                    ids.append(None)
                    logs.append({
                        "row": int(idx) + 1,
                        "result": f"Error: {e}",
                        "contact_id": "",
                        "payload": json.dumps(fields_payload, ensure_ascii=False)
                    })

                progress.progress(min(len(ids)/total, 1.0))

            ok_count = sum(1 for i in ids if i)
            fail_count = total - ok_count
            st.success(f"Done. Success: {ok_count} • Failures: {fail_count}")

            log_df = pd.DataFrame(logs, columns=["row","result","contact_id","payload"])
            st.dataframe(log_df, use_container_width=True)

            # log CSV
            csv_buf = io.StringIO()
            log_df.to_csv(csv_buf, index=False)
            st.download_button(
                "Download log CSV",
                data=csv_buf.getvalue().encode("utf-8"),
                file_name="bitrix_contacts_import_log.csv",
                mime="text/csv"
            )

            # Excel with BITRIX_ID
            try:
                xlsx_bytes = make_excel_with_ids(df, ids)
                st.download_button(
                    "Download Excel with BITRIX_ID",
                    data=xlsx_bytes,
                    file_name="contacts_bitrix_imported.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.info(f"Could not generate output Excel: {e}")

# notes
with st.expander("Tips and notes"):
    st.markdown(
        """
- If the API says a required field is missing, map it before importing.
- You can map multiple columns to EMAIL and PHONE. The app sends them as multi fields.
- Duplicate checker searches by EMAIL and PHONE using crm.contact.list.
- The webhook must allow crm.contact.add.
- Excel dates are converted to ISO (YYYY-MM-DD).
        """
    )
