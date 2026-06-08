import io
import pandas as pd
import msoffcrypto
from getpass import getpass

input_path = "/Users/danielstrick/Downloads/UChicago/PH project/Advanced Hemos PAH Master Database (DS Copy).xlsx"
encrypted_output_path = "/Users/danielstrick/Downloads/UChicago/PH project/patients_first_two_RHCs_wide_ENCRYPTED.xlsx"

password = getpass("Enter Excel password: ")

# ── 1. Decrypt & Load ────────────────────────────────────────────────────────
decrypted_in = io.BytesIO()
with open(input_path, "rb") as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted_in)

decrypted_in.seek(0)
df = pd.read_excel(decrypted_in, sheet_name=0, engine="openpyxl")

# ── 2. Clean & Sort ──────────────────────────────────────────────────────────
df["Date of RHC"] = pd.to_datetime(df["Date of RHC"], errors="coerce")
df = df.dropna(subset=["Date of RHC"])
df = df.drop_duplicates(subset=["Patient ID", "Date of RHC"])
df = df.sort_values(["Patient ID", "Date of RHC"]).reset_index(drop=True)

# ── 3. Resolve SBP/DBP/MAP: Ao preferred, fall back to Cuff ─────────────────
def coalesce(df, ao_col, cuff_col):
    return df[ao_col].where(df[ao_col].notna(), df[cuff_col])

df["SBP"] = coalesce(df, "Ao SBP", "Cuff SBP")
df["DBP"] = coalesce(df, "Ao DBP", "Cuff DBP")
df["MAP"] = coalesce(df, "Ao MAP", "Cuff MAP")

# ── 4. Resolve PCWP/LVEDP: LVEDP preferred, fall back to PCWP ───────────────
df["PCWP_resolved"] = df["LVEDP"].where(df["LVEDP"].notna(), df["PCWP"])

# ── 5. Identify RHC #1 via Initial RHC flag, RHC #2 as next chronological ───
initial_flag_col = "Initial RHC (1 = Y or 0=N)"

# Audit: flag where Initial RHC=1 is not the chronologically earliest visit
chrono_first = df.groupby("Patient ID")["Date of RHC"].transform("min")
discrepancy_mask = (df[initial_flag_col] == 1) & (df["Date of RHC"] != chrono_first)
discrepancy_log = df[discrepancy_mask][["Patient ID", "Date of RHC", initial_flag_col]].copy()
discrepancy_log["Issue"] = "Initial RHC flag=1 but not earliest date"

# Audit: flag patients with no Initial RHC=1 at all
patients_with_flag = df[df[initial_flag_col] == 1]["Patient ID"].unique()
missing_flag_patients = set(df["Patient ID"].unique()) - set(patients_with_flag)
missing_flag_log = pd.DataFrame({
    "Patient ID": list(missing_flag_patients),
    "Date of RHC": [pd.NaT] * len(missing_flag_patients),
    initial_flag_col: [None] * len(missing_flag_patients),
    "Issue": "No Initial RHC flag=1 found for this patient"
})
audit_log = pd.concat([discrepancy_log, missing_flag_log], ignore_index=True)

# RHC #1: trust the flag
rhc1 = df[df[initial_flag_col] == 1].drop_duplicates(subset=["Patient ID"])

# RHC #2: earliest visit strictly after RHC #1
rhc1_dates = rhc1.set_index("Patient ID")["Date of RHC"].rename("rhc1_date")
df2 = df.join(rhc1_dates, on="Patient ID")
df2 = df2[df2["Date of RHC"] > df2["rhc1_date"]]
rhc2 = df2.sort_values(["Patient ID", "Date of RHC"]).drop_duplicates(subset=["Patient ID"])

# Keep only paired patients
paired = set(rhc1["Patient ID"]) & set(rhc2["Patient ID"])
rhc1 = rhc1[rhc1["Patient ID"].isin(paired)].set_index("Patient ID")
rhc2 = rhc2[rhc2["Patient ID"].isin(paired)].set_index("Patient ID")

# ── 6. Build Wide Output using _1/_2 suffixes (no duplicate col names yet) ──
def get_col(frame, col):
    return frame[col] if col in frame.columns else pd.Series(pd.NA, index=frame.index)

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

pid = rhc1.index

cols = {}
cols["Patient ID"]                                   = pid.to_series(index=pid)
cols["Diagnosis"]                                    = rhc1["Diagnosis"]

# RHC #1
cols["Date of RHC #1"]                              = rhc1["Date of RHC"]
cols["SBP_1"]                                        = rhc1["SBP"]
cols["DBP_1"]                                        = rhc1["DBP"]
cols["MAP_1"]                                        = rhc1["MAP"]
cols["HR_1"]                                         = get_col(rhc1, "HR")
cols["SpO2_1"]                                       = pd.Series(pd.NA, index=pid)  # not in source
cols["RA_1"]                                         = get_col(rhc1, "RA")
cols["RVSP_1"]                                       = get_col(rhc1, "RVSP")
cols["RVDP_1"]                                       = get_col(rhc1, "RVDP")
cols["PASP_1"]                                       = get_col(rhc1, "PASP")
cols["PADP_1"]                                       = get_col(rhc1, "PADP")
cols["MPAP_1"]                                       = get_col(rhc1, "MPAP")
cols["PCWP or LVEDP (LVEDP n=288, PCW =268)_1"]     = rhc1["PCWP_resolved"]
cols["PA Sat_1"]                                     = get_col(rhc1, "PA Sat")
cols["Stroke Vol_1"]                                 = get_col(rhc1, "Stroke Vol")
cols["Stroke Vol Index_1"]                           = get_col(rhc1, "Stroke Vol Index")
cols["PAPP_1"]                                       = get_col(rhc1, "PAPP")
cols["Cpa (TD)_1"]                                   = get_col(rhc1, "Cpa (TD)")
cols["TD CO_1"]                                      = get_col(rhc1, "TD CO")
cols["TD CI_1"]                                      = get_col(rhc1, "TD CI")
cols["TD PVR_1"]                                     = get_col(rhc1, "TD PVR")
cols["TD SVR_1"]                                     = get_col(rhc1, "TD SVR")
cols["PAPI_1"]                                       = get_col(rhc1, "PAPI")
cols["RV-CPO_1"]                                     = get_col(rhc1, "RV-CPO")
cols["RV-MPS_1"]                                     = get_col(rhc1, "RV-MPS")

# RHC #2
cols["Date of RHC #2"]                              = rhc2["Date of RHC"]
cols["PDE5i_2"]                                      = pd.Series(pd.NA, index=pid)  # placeholder
cols["SGC_2"]                                        = pd.Series(pd.NA, index=pid)
cols["ERA_2"]                                        = pd.Series(pd.NA, index=pid)
cols["PRA_2"]                                        = pd.Series(pd.NA, index=pid)
cols["INH PCA_2"]                                    = pd.Series(pd.NA, index=pid)
cols["SQ PCA_2"]                                     = pd.Series(pd.NA, index=pid)
cols["IV PCA_2"]                                     = pd.Series(pd.NA, index=pid)
cols["Total # of Agents at time of RHC #2_2"]       = pd.Series(pd.NA, index=pid)
cols["BSA_2"]                                        = pd.Series(pd.NA, index=pid)
cols["SBP_2"]                                        = rhc2["SBP"]
cols["DBP_2"]                                        = rhc2["DBP"]
cols["MAP_2"]                                        = rhc2["MAP"]
cols["HR_2"]                                         = get_col(rhc2, "HR")
cols["RA_2"]                                         = get_col(rhc2, "RA")
cols["RVSP_2"]                                       = get_col(rhc2, "RVSP")
cols["RVDP_2"]                                       = get_col(rhc2, "RVDP")
cols["PASP_2"]                                       = get_col(rhc2, "PASP")
cols["PADP_2"]                                       = get_col(rhc2, "PADP")
cols["MPAP_2"]                                       = get_col(rhc2, "MPAP")
cols["PCWP or LVEDP (LVEDP preferred)_2"]            = rhc2["PCWP_resolved"]
cols["PA Sat_2"]                                     = get_col(rhc2, "PA Sat")
cols["Fick CO_2"]                                    = get_col(rhc2, "Fick CO")
cols["Fick CI_2"]                                    = get_col(rhc2, "Fick CI")
cols["TD CO_2"]                                      = get_col(rhc2, "TD CO")
cols["TD CI_2"]                                      = get_col(rhc2, "TD CI")
cols["Fick PVR_2"]                                   = get_col(rhc2, "Fick PVR")
cols["TD PVR_2"]                                     = get_col(rhc2, "TD PVR")
cols["Fick SVR_2"]                                   = get_col(rhc2, "Fick SVR")
cols["TD SVR_2"]                                     = get_col(rhc2, "TD SVR")
cols["Stroke Vol_2"]                                 = get_col(rhc2, "Stroke Vol")
cols["Stroke Vol Index_2"]                           = get_col(rhc2, "Stroke Vol Index")
cols["PAPP_2"]                                       = get_col(rhc2, "PAPP")
cols["Cpa_2"]                                        = get_col(rhc2, "Cpa")
cols["PAPI_2"]                                       = get_col(rhc2, "PAPI")
cols["RV-CPO_2"]                                     = get_col(rhc2, "RV-CPO")
cols["RV-MPS_2"]                                     = get_col(rhc2, "RV-MPS")

base = pd.DataFrame(cols)

# ── 7. Delta Columns ─────────────────────────────────────────────────────────
base["∆MPAP"]   = to_num(base["MPAP_2"])            - to_num(base["MPAP_1"])
base["∆PA Sat"] = to_num(base["PA Sat_2"])          - to_num(base["PA Sat_1"])
base["∆SVi"]    = to_num(base["Stroke Vol Index_2"])- to_num(base["Stroke Vol Index_1"])
base["∆Cpa"]    = to_num(base["Cpa_2"])             - to_num(base["Cpa (TD)_1"])
base["∆PVR"]    = to_num(base["TD PVR_2"])          - to_num(base["TD PVR_1"])
base["∆PAPI"]   = to_num(base["PAPI_2"])            - to_num(base["PAPI_1"])
base["∆RV-CPO"] = to_num(base["RV-CPO_2"])          - to_num(base["RV-CPO_1"])
base["∆RV-MPS"] = to_num(base["RV-MPS_2"])          - to_num(base["RV-MPS_1"])

# ── 8. Strip _1/_2 suffixes for final output headers ─────────────────────────
# removesuffix is Python 3.9+; the _1 pass then _2 pass handles both
final_df = base.rename(columns=lambda c: c.removesuffix("_1").removesuffix("_2"))

# ── 9. Validation ─────────────────────────────────────────────────────────────
print(f"Patients with paired RHCs: {len(final_df)}")
all_null = [c for i, c in enumerate(final_df.columns) if final_df.iloc[:, i].isna().all()]
non_placeholder_null = [c for c in all_null if c not in {
    "SpO2", "PDE5i", "SGC", "ERA", "PRA", "INH PCA", "SQ PCA",
    "IV PCA", "Total # of Agents at time of RHC #2", "BSA"
}]
if non_placeholder_null:
    print(f"WARNING — unexpected all-null columns (check source col names): {non_placeholder_null}")
if not audit_log.empty:
    print(f"AUDIT — {len(audit_log)} discrepancies logged to 'Audit' sheet")

# ── 10. Encrypt & Save ────────────────────────────────────────────────────────
unencrypted_out = io.BytesIO()
with pd.ExcelWriter(unencrypted_out, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Wide Data", index=False)
    if not audit_log.empty:
        audit_log.to_excel(writer, sheet_name="Audit", index=False)

unencrypted_out.seek(0)
encrypted_out = io.BytesIO()
file_to_encrypt = msoffcrypto.OfficeFile(unencrypted_out)
file_to_encrypt.load_key(password=password)
file_to_encrypt.encrypt(password, encrypted_out)

with open(encrypted_output_path, "wb") as f:
    f.write(encrypted_out.getvalue())

del password
print(f"Done. Saved to: {encrypted_output_path}")