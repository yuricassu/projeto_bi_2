# app.py
import streamlit as st
import pandas as pd
import zipfile, json, re
from io import BytesIO

st.set_page_config(page_title="Auditoria Power BI", layout="wide")

st.title("🔎 Auditoria de Modelos Power BI (.pbit)")

# -----------------------------
# Upload do arquivo .pbit
# -----------------------------
uploaded_file = st.file_uploader("Escolha o arquivo .pbit", type="pbit")

if uploaded_file:
    st.success(f"Arquivo carregado: {uploaded_file.name}")

    # -----------------------------
    # Funções auxiliares
    # -----------------------------
    DAX_REF_PATTERN = re.compile(r"'?([A-Za-z0-9_ ]+)'?\[([A-Za-z0-9_ ]+)\]")

    def extract_table_column_refs_from_text(text):
        used = set()
        if not text:
            return used
        if isinstance(text, list):
            text = "\n".join(text)
        if not isinstance(text, str):
            return used
        for m in DAX_REF_PATTERN.finditer(text):
            if m.group(1) and m.group(2):
                used.add((m.group(1).strip(), m.group(2).strip()))
        return used

    # -----------------------------
    # Função de auditoria
    # -----------------------------
    def audit_model(pbit_file_obj):
        results = {
            "unused_columns": [],
            "duplicate_measures": [],
            "missing_descriptions": [],
            "orphan_tables": []
        }

        with zipfile.ZipFile(pbit_file_obj, "r") as z:
            model_json = json.loads(z.read("DataModelSchema"))

        tables = model_json.get("model", {}).get("tables", [])
        relationships = model_json.get("model", {}).get("relationships", [])

        measure_list = []
        used_in_measures = set()

        for t in tables:
            tname = t["name"]
            for c in t.get("columns", []):
                cname = c["name"]
                desc = c.get("description", "")
                if not desc:
                    results["missing_descriptions"].append((tname, cname))
            for m in t.get("measures", []):
                mname = m["name"]
                expr = m.get("expression", "") or ""
                if isinstance(expr, list):
                    expr = "\n".join(expr)
                desc = m.get("description", "")
                measure_list.append({"table": tname, "measure": mname, "expression": expr, "desc": desc})
                used_in_measures |= extract_table_column_refs_from_text(expr)
                if not desc:
                    results["missing_descriptions"].append((tname, mname))

        for t in tables:
            tname = t["name"]
            for c in t.get("columns", []):
                cname = c["name"]
                from_list = [(r.get("fromTable"), r.get("fromColumn")) for r in relationships]
                to_list = [(r.get("toTable"), r.get("toColumn")) for r in relationships]
                if (tname, cname) not in used_in_measures and \
                   (tname, cname) not in from_list and \
                   (tname, cname) not in to_list:
                    results["unused_columns"].append((tname, cname))

        expr_map = {}
        for m in measure_list:
            expr = m["expression"] or ""
            norm_expr = expr.strip().replace(" ", "").lower()
            if norm_expr in expr_map:
                results["duplicate_measures"].append((expr_map[norm_expr], m))
            else:
                expr_map[norm_expr] = m

        related_tables = set([r["fromTable"] for r in relationships] + [r["toTable"] for r in relationships])
        for t in tables:
            if t["name"] not in related_tables:
                results["orphan_tables"].append(t["name"])

        results_df = {
            "unused_columns": pd.DataFrame(results["unused_columns"], columns=["table", "column"]),
            "duplicate_measures": pd.DataFrame([
                {"table1": m1["table"], "measure1": m1["measure"], 
                 "table2": m2["table"], "measure2": m2["measure"], 
                 "expression": m1["expression"]}
                for m1, m2 in results["duplicate_measures"]
            ]),
            "missing_descriptions": pd.DataFrame(results["missing_descriptions"], columns=["table", "name"]),
            "orphan_tables": pd.DataFrame(results["orphan_tables"], columns=["table"])
        }

        return results_df, tables, measure_list

    # -----------------------------
    # Rodar auditoria
    # -----------------------------
    results, all_tables, all_measures = audit_model(uploaded_file)

    # -----------------------------
    # Mostrar resumo
    # -----------------------------
    st.header("📊 Resumo da Auditoria")
    for key, df in results.items():
        st.subheader(key.replace("_", " ").title())
        if df.empty:
            st.info("Nenhum problema encontrado")
        else:
            st.dataframe(df)

    # -----------------------------
    # Ranking visual
    # -----------------------------
    st.header("🏆 Ranking de Problemas por Tabela")
    ranking_df = pd.DataFrame({"table": [t["name"] for t in all_tables]})
    ranking_df["colunas_nao_usadas"] = ranking_df["table"].apply(
        lambda x: len([c for t,c in results["unused_columns"].values if t==x])
    )
    ranking_df["campos_sem_descricao"] = ranking_df["table"].apply(
        lambda x: len([c for t,c in results["missing_descriptions"].values if t==x])
    )
    if results["duplicate_measures"].empty:
        ranking_df["medidas_duplicadas"] = 0
    else:
        ranking_df["medidas_duplicadas"] = ranking_df["table"].apply(
            lambda x: len([m for m in results["duplicate_measures"]["table1"].values if m==x])
        )
    st.dataframe(ranking_df)

    # -----------------------------
    # Download Excel
    # -----------------------------
    st.header("💾 Download Excel Completo")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        ranking_df.to_excel(writer, sheet_name="Ranking_Problemas", index=False)
    output.seek(0)
    st.download_button("📥 Baixar Excel da Auditoria", data=output, file_name="Auditoria_Modelo_PBI.xlsx")
