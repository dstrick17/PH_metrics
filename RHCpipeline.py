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
def coalesce(a, b):
    return a.where(a.notna(), b)

df["SBP"] = coalesce(df["Ao SBP"], df["Cuff SBP"])
df["DBP"] = coalesce(df["Ao DBP"], df["Cuff DBP"])
df["MAP"] = coalesce(df["Ao MAP"], df["Cuff MAP"])

# ── 4. Resolve PCWP/LVEDP: LVEDP preferred, fall back to PCWP ───────────────
df["PCWP_resolved"] = coalesce(df["LVEDP"], df["PCWP"])

# ── 5. Identify RHC #1 via Initial RHC flag, RHC #2 as next chronological ───
initial_flag_col = "Initial RHC (1 = Y or 0=N)"

# Audit: Initial RHC=1 flag is not the chronologically earliest visit
chrono_first = df.groupby("Patient ID")["Date of RHC"].transform("min")
discrepancy_mask = (df[initial_flag_col] == 1) & (df["Date of RHC"] != chrono_first)
discrepancy_log = df[discrepancy_mask][["Patient ID", "Date of RHC", initial_flag_col]].copy()
discrepancy_log["Issue"] = "Initial RHC flag=1 but not earliest date"

# Audit: patients with no Initial RHC=1 at all
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

# ── 6. Helpers ───────────────────────────────────────────────────────────────
def get(frame, col):
    """Safely get a column; return all-NA series if column doesn't exist."""
    return frame[col] if col in frame.columns else pd.Series(pd.NA, index=frame.index)

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def blank(frame):
    return pd.Series(pd.NA, index=frame.index)

pid = rhc1.index

# ── 7. Build Wide Output ─────────────────────────────────────────────────────
cols = {}

# ── Patient-level metadata (from RHC #1 row) ─────────────────────────────────
cols["Patient ID"]                              = pid.to_series(index=pid)
cols["Diagnosis"]                               = get(rhc1, "Diagnosis")
cols["Age"]                                     = get(rhc1, "Age")
cols["Sex (1=M or 0=F)"]                        = get(rhc1, "Sex (1=M or 0=F)")
cols["Race (1=White, 2=Black, 3=Hispanic, 4=Asian, 5=Other)"] = get(rhc1, "Race (1=White, 2=Black, 3=Hispanic, 4=Asian, 5=Other)")
cols["WHO FC"]                                  = get(rhc1, "WHO FC")
cols["NT-ProBNP"]                               = get(rhc1, "NT-ProBNP")
cols["eGFR"]                                    = get(rhc1, "eGFR")
cols["6MWD (meters)"]                           = get(rhc1, "6MWD")

# ── Outcome columns (patient-level) ──────────────────────────────────────────
cols["Date of RHC #1"]                          = get(rhc1, "Date of RHC")
cols["Date of Death"]                           = get(rhc1, "Date of Death")
cols["Date of Transplant"]                      = get(rhc1, "Date of Transplant")
cols["Death or LTx (1 = Y or 0 = N)"]          = get(rhc1, "Death or LTx (1 = Y or 0 = N)")
cols["Time to Death/LTx (days)"]               = get(rhc1, "Time to Death/LTx (days)")
cols["Time to Death/LTx (years)"]              = get(rhc1, "Time to Death/LTx (years)")
cols["Death or LTx at 1 year (1=Y)"]           = get(rhc1, "Death or LTx at 1 year (1=Y)")
cols["Alive w/o LTx in 1/1/2025 (1=Y)"]        = get(rhc1, "Alive in 1/1/2025 (1=Y)")
cols["Total Follow Up Days (censored)"]         = blank(rhc1)   # not in source

# ── RHC #1 hemodynamics ───────────────────────────────────────────────────────
cols["BSA_1"]                                   = blank(rhc1)   # placeholder
cols["SBP_1"]                                   = rhc1["SBP"]
cols["DBP_1"]                                   = rhc1["DBP"]
cols["MAP_1"]                                   = rhc1["MAP"]
cols["HR_1"]                                    = get(rhc1, "HR")
cols["SpO2_1"]                                  = get(rhc1, "SpO2")
cols["RA_1"]                                    = get(rhc1, "RA")
cols["RVSP_1"]                                  = get(rhc1, "RVSP")
cols["RVDP_1"]                                  = get(rhc1, "RVDP")
cols["PASP_1"]                                  = get(rhc1, "PASP")
cols["PADP_1"]                                  = get(rhc1, "PADP")
cols["MPAP_1"]                                  = get(rhc1, "MPAP")
cols["PCWP or LVEDP (LVEDP n=288, PCW =268)_1"]= rhc1["PCWP_resolved"]
cols["PA Sat_1"]                                = get(rhc1, "PA Sat")
cols["Stroke Vol_1"]                            = blank(rhc1)   # PI to calculate
cols["Stroke Vol Index_1"]                      = blank(rhc1)
cols["PAPP_1"]                                  = blank(rhc1)
cols["Cpa (TD)_1"]                              = blank(rhc1)
cols["TD CO_1"]                                 = get(rhc1, "TD CO")
cols["TD CI_1"]                                 = get(rhc1, "TD CI")
cols["TD PVR_1"]                                = get(rhc1, "TD PVR")
cols["TD SVR_1"]                                = get(rhc1, "TD SVR")
cols["PAPI_1"]                                  = get(rhc1, "PAPI")
cols["RV-CPO (Fick)_1"]                         = get(rhc1, "RV-CPO (Fick)")
cols["RV-CPO (TD)_1"]                           = get(rhc1, "RV-CPO (TD)")
cols["RV-MPS (Fick)_1"]                         = get(rhc1, "RV-MPS (Fick)")
cols["RV-MPS (TD)_1"]                           = get(rhc1, "RV-MPS (TD)")

# ── RHC #2 hemodynamics ───────────────────────────────────────────────────────
cols["Date of RHC #2"]                          = get(rhc2, "Date of RHC")
cols["PDE5i_2"]                                 = blank(rhc2)   # placeholder
cols["SGC_2"]                                   = blank(rhc2)
cols["ERA_2"]                                   = blank(rhc2)
cols["PRA_2"]                                   = blank(rhc2)
cols["INH PCA_2"]                               = blank(rhc2)
cols["SQ PCA_2"]                                = blank(rhc2)
cols["IV PCA_2"]                                = blank(rhc2)
cols["Total # of Agents at time of RHC #2_2"]  = blank(rhc2)
cols["BSA_2"]                                   = blank(rhc2)   # placeholder
cols["SBP_2"]                                   = rhc2["SBP"]
cols["DBP_2"]                                   = rhc2["DBP"]
cols["MAP_2"]                                   = rhc2["MAP"]
cols["HR_2"]                                    = get(rhc2, "HR")
cols["RA_2"]                                    = get(rhc2, "RA")
cols["RVSP_2"]                                  = get(rhc2, "RVSP")
cols["RVDP_2"]                                  = get(rhc2, "RVDP")
cols["PASP_2"]                                  = get(rhc2, "PASP")
cols["PADP_2"]                                  = get(rhc2, "PADP")
cols["MPAP_2"]                                  = get(rhc2, "MPAP")
cols["PCWP or LVEDP (LVEDP preferred)_2"]       = rhc2["PCWP_resolved"]
cols["PA Sat_2"]                                = get(rhc2, "PA Sat")
cols["Fick CO_2"]                               = get(rhc2, "Fick CO")
cols["Fick CI_2"]                               = get(rhc2, "Fick CI")
cols["TD CO_2"]                                 = get(rhc2, "TD CO")
cols["TD CI_2"]                                 = get(rhc2, "TD CI")
cols["Fick PVR_2"]                              = get(rhc2, "Fick PVR")
cols["TD PVR_2"]                                = get(rhc2, "TD PVR")
cols["Fick SVR_2"]                              = get(rhc2, "Fick SVR")
cols["TD SVR_2"]                                = get(rhc2, "TD SVR")
cols["Stroke Vol_2"]                            = blank(rhc2)   # PI to calculate
cols["Stroke Vol Index_2"]                      = blank(rhc2)
cols["PAPP_2"]                                  = blank(rhc2)
cols["Cpa_2"]                                   = blank(rhc2)
cols["PAPI_2"]                                  = get(rhc2, "PAPI")
cols["RV-CPO (Fick)_2"]                         = get(rhc2, "RV-CPO (Fick)")
cols["RV-CPO (TD)_2"]                           = get(rhc2, "RV-CPO (TD)")
cols["RV-MPS (Fick)_2"]                         = get(rhc2, "RV-MPS (Fick)")
cols["RV-MPS (TD)_2"]                           = get(rhc2, "RV-MPS (TD)")

# ── Vasoreactivity columns (as-is from source, no suffix stripping) ───────────
cols["Agent"]                                   = get(rhc2, "Agent")
cols["Dose"]                                    = get(rhc2, "Dose")
cols["RA 2"]                                    = get(rhc2, "RA 2")
cols["RVSP 2"]                                  = get(rhc2, "RVSP 2")
cols["RVDP 2"]                                  = get(rhc2, "RVDP 2")
cols["PASP 2"]                                  = get(rhc2, "PASP 2")
cols["PADP 2"]                                  = get(rhc2, "PADP 2")
cols["MPAP 2"]                                  = get(rhc2, "MPAP 2")
cols["PCWP 2"]                                  = get(rhc2, "PCWP 2")
cols["LVEDP 2"]                                 = get(rhc2, "LVEDP 2")
cols["PA Sat 2"]                                = get(rhc2, "PA Sat 2")
cols["Fick CO 2"]                               = get(rhc2, "Fick CO 2")
cols["Fick CI 2"]                               = get(rhc2, "Fick CI 2")
cols["TD CO 2"]                                 = get(rhc2, "TD CO 2")
cols["TD CI 2"]                                 = get(rhc2, "TD CI 2")
cols["Fick PVR 2"]                              = get(rhc2, "Fick PVR 2")
cols["TD PVR 2"]                                = get(rhc2, "TD PVR 2")

# ── Delta columns ─────────────────────────────────────────────────────────────
base = pd.DataFrame(cols)

base["∆MPAP"]   = to_num(base["MPAP_2"])             - to_num(base["MPAP_1"])
base["∆PA Sat"] = to_num(base["PA Sat_2"])           - to_num(base["PA Sat_1"])
base["∆SVi"]    = to_num(base["Stroke Vol Index_2"]) - to_num(base["Stroke Vol Index_1"])
base["∆Cpa"]    = to_num(base["Cpa_2"])              - to_num(base["Cpa (TD)_1"])
base["∆PVR"]    = to_num(base["TD PVR_2"])           - to_num(base["TD PVR_1"])
base["∆PAPI"]   = to_num(base["PAPI_2"])             - to_num(base["PAPI_1"])
base["∆RV-CPO (Fick)"] = to_num(base["RV-CPO (Fick)_2"]) - to_num(base["RV-CPO (Fick)_1"])
base["∆RV-CPO (TD)"]   = to_num(base["RV-CPO (TD)_2"])   - to_num(base["RV-CPO (TD)_1"])
base["∆RV-MPS (Fick)"] = to_num(base["RV-MPS (Fick)_2"]) - to_num(base["RV-MPS (Fick)_1"])
base["∆RV-MPS (TD)"]   = to_num(base["RV-MPS (TD)_2"])   - to_num(base["RV-MPS (TD)_1"])

# ── Strip _1/_2 suffixes (vasoreactivity cols and delta cols are unaffected) ──
KEEP_SUFFIX = {   # these end in a real space+digit, not our suffix — leave alone
    "RA 2", "RVSP 2", "RVDP 2", "PASP 2", "PADP 2", "MPAP 2",
    "PCWP 2", "LVEDP 2", "PA Sat 2",
    "Fick CO 2", "Fick CI 2", "TD CO 2", "TD CI 2", "Fick PVR 2", "TD PVR 2"
}

def strip_suffix(c):
    if c in KEEP_SUFFIX:
        return c
    return c.removesuffix("_1").removesuffix("_2")

final_df = base.rename(columns=strip_suffix)

# ── 8. Validation ─────────────────────────────────────────────────────────────
KNOWN_PLACEHOLDERS = {
    "BSA", "Total Follow Up Days (censored)",
    "Stroke Vol", "Stroke Vol Index", "PAPP", "Cpa (TD)", "Cpa",
    "PDE5i", "SGC", "ERA", "PRA", "INH PCA", "SQ PCA", "IV PCA",
    "Total # of Agents at time of RHC #2",
}

print(f"Patients with paired RHCs: {len(final_df)}")
all_null = [c for i, c in enumerate(final_df.columns) if final_df.iloc[:, i].isna().all()]
unexpected_null = [c for c in all_null if c not in KNOWN_PLACEHOLDERS]
if unexpected_null:
    print(f"WARNING — unexpected all-null columns (check source col names): {unexpected_null}")
else:
    print("All non-placeholder columns populated.")
if not audit_log.empty:
    print(f"AUDIT — {len(audit_log)} discrepancies logged to 'Audit' sheet")

# ── 9. Encrypt & Save ─────────────────────────────────────────────────────────
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