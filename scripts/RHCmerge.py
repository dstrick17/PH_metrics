import io
import pandas as pd
import msoffcrypto
from getpass import getpass
from pathlib import Path

project_dir = Path(__file__).resolve().parent.parent
master_path = project_dir / "input files" / "Advanced Hemos PAH Master Database (DS Copy).xlsx"
pi_path = project_dir / "input files" / "Sequential RV-MPS in PAH (DS Copy).xlsx"
output_path = project_dir / "output files" / "Sequential RV-MPS in PAH (DS Copy - Updated).xlsx"

password = getpass("Enter Excel password: ")

# ── Helpers ──────────────────────────────────────────────────────────────────
def decrypt_file(path, password):
    buf = io.BytesIO()
    with open(path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password=password)
        office.decrypt(buf)
    buf.seek(0)
    return buf

def coalesce(a, b):
    return a.where(a.notna(), b)

def first_available(row, cols):
    for col in cols:
        val = row.get(col, pd.NA)
        if pd.notna(val):
            return val
    return pd.NA

# ── 1. Load both files ────────────────────────────────────────────────────────
print("Loading master database...")
master_buf = decrypt_file(master_path, password)
master = pd.read_excel(master_buf, sheet_name=0, engine="openpyxl")

print("Loading PI-edited spreadsheet...")
pi_buf = decrypt_file(pi_path, password)
pi_df  = pd.read_excel(pi_buf, sheet_name=0, engine="openpyxl")

# ── 2. Prepare master ────────────────────────────────────────────────────────
master["Date of RHC"] = pd.to_datetime(master["Date of RHC"], errors="coerce")
master = master.dropna(subset=["Date of RHC"])
master = master.drop_duplicates(subset=["Patient ID", "Date of RHC"])
master = master.sort_values(["Patient ID", "Date of RHC"]).reset_index(drop=True)

master["SBP"] = coalesce(master["Ao SBP"], master["Cuff SBP"])
master["DBP"] = coalesce(master["Ao DBP"], master["Cuff DBP"])
master["MAP"] = coalesce(master["Ao MAP"], master["Cuff MAP"])
master["PCWP_resolved"] = coalesce(master["LVEDP"], master["PCWP"])
master["CO_resolved"] = coalesce(master["TD CO"], master["Fick CO"])
master["CI_resolved"] = coalesce(master["TD CI"], master["Fick CI"])
master["PVR_resolved"] = coalesce(master["TD PVR"], master["Fick PVR"])
master["SVR_resolved"] = coalesce(master["TD SVR"], master["Fick SVR"])
master["RV-CPO_resolved"] = coalesce(master["RV-CPO (TD)"], master["RV-CPO (Fick)"])
master["RV-MPS_resolved"] = coalesce(master["RV-MPS (TD)"], master["RV-MPS (Fick)"])

initial_flag = "Initial RHC (1 = Y or 0=N)"

# RHC #1: trust the Initial RHC flag
rhc1 = master[master[initial_flag] == 1].drop_duplicates(subset=["Patient ID"])

# RHC #2: earliest visit strictly after RHC #1
rhc1_dates = rhc1.set_index("Patient ID")["Date of RHC"].rename("rhc1_date")
df2 = master.join(rhc1_dates, on="Patient ID")
df2 = df2[df2["Date of RHC"] > df2["rhc1_date"]]
rhc2 = df2.sort_values(["Patient ID", "Date of RHC"]).drop_duplicates(subset=["Patient ID"])

# Keep only patients with both RHC #1 and RHC #2
paired = set(rhc1["Patient ID"]) & set(rhc2["Patient ID"])
rhc1 = rhc1[rhc1["Patient ID"].isin(paired)].set_index("Patient ID")
rhc2 = rhc2[rhc2["Patient ID"].isin(paired)].set_index("Patient ID")

print(f"Master: {len(paired)} patients with paired RHCs found")

# ── 3. Define PI sheet column layout by position ──────────────────────────────
# Column indices (0-based) for the PI sheet header layout:
# Patient ID=0, Diagnosis=1, Age=2, Sex=3, Race=4, WHO FC=5, NT-ProBNP=6,
# eGFR=7, 6MWD=8,
# --- RHC #1 block ---
# Date of RHC #1=9, Date of Death=10, Date of Transplant=11,
# Death or LTx=12, Time to Death/LTx (days)=13, Time to Death/LTx (years)=14,
# Death or LTx at 1 year=15, Alive w/o LTx=16, Total Follow Up Days=17,
# BSA=18, SBP=19, DBP=20, MAP=21, HR=22, SpO2=23, RA=24, RVSP=25, RVDP=26,
# PASP=27, PADP=28, MPAP=29, PCWP or LVEDP=30, PA Sat=31,
# Stroke Vol=32, Stroke Vol Index=33, PAPP=34, Cpa (TD)=35,
# TD CO=36, TD CI=37, TD PVR=38, TD SVR=39, PAPI=40, RV-CPO=41, RV-MPS=42,
# --- RHC #2 block ---
# Date of RHC #2=43, PDE5i=44, SGC=45, ERA=46, PRA=47, INH PCA=48,
# SQ PCA=49, IV PCA=50, Total # of Agents=51, BSA=52,
# SBP=53, DBP=54, MAP=55, HR=56, RA=57, RVSP=58, RVDP=59,
# PASP=60, PADP=61, MPAP=62, PCWP or LVEDP=63, PA Sat=64,
# Fick CO=65, Fick CI=66, TD CO=67, TD CI=68, Fick PVR=69, TD PVR=70,
# Fick SVR=71, TD SVR=72, Stroke Vol=73, Stroke Vol Index=74, PAPP=75,
# Cpa=76, PAPI=77, RV-CPO=78, RV-MPS=79,
# --- Delta block ---
# ∆MPAP=80, ∆PA Sat=81, ∆SVi=82, ∆Cpa=83, ∆PVR=84, ∆PAPI=85,
# ∆RV-CPO=86, ∆RV-MPS=87

# Map: PI column index -> (master_frame, master_column)
# Only columns sourced directly from master (not PI-entered, not calculated)
# For RHC #2 block, master_frame = rhc2
RHC2_MAP = {
    43:  ("rhc2", "Date of RHC"),
    # 44-52: PI-entered placeholders — skip
    53:  ("rhc2", "SBP"),
    54:  ("rhc2", "DBP"),
    55:  ("rhc2", "MAP"),
    56:  ("rhc2", "HR"),
    57:  ("rhc2", "RA"),
    58:  ("rhc2", "RVSP"),
    59:  ("rhc2", "RVDP"),
    60:  ("rhc2", "PASP"),
    61:  ("rhc2", "PADP"),
    62:  ("rhc2", "MPAP"),
    63:  ("rhc2", "PCWP_resolved"),
    64:  ("rhc2", "PA Sat"),
    65:  ("rhc2", "Fick CO"),
    66:  ("rhc2", "Fick CI"),
    67:  ("rhc2", "TD CO"),
    68:  ("rhc2", "TD CI"),
    69:  ("rhc2", "Fick PVR"),
    70:  ("rhc2", "TD PVR"),
    71:  ("rhc2", "Fick SVR"),
    72:  ("rhc2", "TD SVR"),
    # 73-76: Stroke Vol, Stroke Vol Index, PAPP, Cpa — PI to calculate, skip
    77:  ("rhc2", "PAPI"),
    78:  ("rhc2", "RV-CPO_resolved"),
    79:  ("rhc2", "RV-MPS_resolved"),
}

# Delta columns computed from the PI sheet itself after RHC2 is filled
# col_idx: ([rhc2_col_idx choices], [rhc1_col_idx choices])
DELTA_MAP = {
    80: ([62], [29]),       # ∆MPAP   = MPAP_2   - MPAP_1
    81: ([64], [31]),       # ∆PA Sat = PA Sat_2 - PA Sat_1
    82: ([74], [33]),       # ∆SVi    = SVI_2    - SVI_1
    83: ([76], [35]),       # ∆Cpa    = Cpa_2    - Cpa_1
    84: ([70, 69], [38]),   # ∆PVR    = TD PVR_2 if present, else Fick PVR_2
    85: ([77], [40]),       # ∆PAPI   = PAPI_2   - PAPI_1
    86: ([78], [41]),       # ∆RV-CPO = resolved RV-CPO_2 - resolved RV-CPO_1
    87: ([79], [42]),       # ∆RV-MPS = resolved RV-MPS_2 - resolved RV-MPS_1
}

# Optional delta columns. These are calculated only if matching PI-sheet headers exist.
MASTER_NAMED_DELTA_MAP = {
    "∆CO": (["CO_resolved"], ["CO_resolved"]),
    "∆CI": (["CI_resolved"], ["CI_resolved"]),
    "∆SVR": (["SVR_resolved"], ["SVR_resolved"]),
}

# ── 4. Filter PI sheet to patients that have a paired RHC in master ───────────
pi_patient_col = pi_df.columns[0]   # Patient ID is col 0
pi_df[pi_patient_col] = pi_df[pi_patient_col].astype(str).str.strip()
rhc1.index = rhc1.index.astype(str).str.strip()
rhc2.index = rhc2.index.astype(str).str.strip()

patients_in_master = set(rhc2.index)
pi_filtered = pi_df[pi_df[pi_patient_col].isin(patients_in_master)].copy().reset_index(drop=True)

print(f"PI sheet: {len(pi_df)} patients total, "
      f"{len(pi_filtered)} have a matching paired RHC in master")

# ── 5. Fill RHC #2 columns — skip if PI already has a value ──────────────────
filled_counts  = {i: 0 for i in RHC2_MAP}
skipped_counts = {i: 0 for i in RHC2_MAP}

for row_idx, row in pi_filtered.iterrows():
    pid = str(row[pi_patient_col]).strip()
    if pid not in rhc2.index:
        continue
    master_row = rhc2.loc[pid]

    for col_idx, (_, master_col) in RHC2_MAP.items():
        pi_col_name = pi_filtered.columns[col_idx]
        current_val = pi_filtered.at[row_idx, pi_col_name]
        master_val  = master_row.get(master_col, pd.NA) if isinstance(master_row, pd.Series) else pd.NA

        # Skip if PI already has a value
        if pd.notna(current_val):
            skipped_counts[col_idx] += 1
            continue

        # Only fill if master has a value
        if pd.notna(master_val):
            pi_filtered.at[row_idx, pi_col_name] = master_val
            filled_counts[col_idx] += 1

# ── 6. Compute delta columns (skip if PI already has a value) ────────────────
def to_num(s): return pd.to_numeric(s, errors="coerce")

def numeric_first_from_row(row, cols):
    return to_num(pd.Series([first_available(row, cols)])).iloc[0]

def numeric_first_from_pi(row_idx, col_indices):
    for col_idx in col_indices:
        val = to_num(pd.Series([pi_filtered.iloc[row_idx, col_idx]])).iloc[0]
        if pd.notna(val):
            return val
    return pd.NA

MASTER_DELTA_MAP = {
    84: (["PVR_resolved"], ["PVR_resolved"]),
    86: (["RV-CPO_resolved"], ["RV-CPO_resolved"]),
    87: (["RV-MPS_resolved"], ["RV-MPS_resolved"]),
}

for row_idx, row in pi_filtered.iterrows():
    pid = str(row[pi_patient_col]).strip()
    if pid not in rhc1.index or pid not in rhc2.index:
        continue

    rhc1_row = rhc1.loc[pid]
    rhc2_row = rhc2.loc[pid]

    for delta_idx, (rhc2_cols, rhc1_cols) in MASTER_DELTA_MAP.items():
        if delta_idx >= len(pi_filtered.columns):
            continue
        if pd.notna(pi_filtered.iloc[row_idx, delta_idx]):
            continue
        v2 = numeric_first_from_row(rhc2_row, rhc2_cols)
        v1 = numeric_first_from_row(rhc1_row, rhc1_cols)
        if pd.notna(v2) and pd.notna(v1):
            pi_filtered.iloc[row_idx, delta_idx] = v2 - v1

    for delta_name, (rhc2_cols, rhc1_cols) in MASTER_NAMED_DELTA_MAP.items():
        if delta_name not in pi_filtered.columns:
            continue
        delta_idx = pi_filtered.columns.get_loc(delta_name)
        if pd.notna(pi_filtered.iloc[row_idx, delta_idx]):
            continue
        v2 = numeric_first_from_row(rhc2_row, rhc2_cols)
        v1 = numeric_first_from_row(rhc1_row, rhc1_cols)
        if pd.notna(v2) and pd.notna(v1):
            pi_filtered.iloc[row_idx, delta_idx] = v2 - v1

for delta_idx, (rhc2_indices, rhc1_indices) in DELTA_MAP.items():
    delta_col  = pi_filtered.columns[delta_idx]
    rhc2_col   = pi_filtered.columns[rhc2_indices[0]]
    rhc1_col   = pi_filtered.columns[rhc1_indices[0]]

    for row_idx in range(len(pi_filtered)):
        current = pi_filtered.iloc[row_idx, delta_idx]
        if pd.notna(current):
            continue   # PI already has a value — skip
        v2 = numeric_first_from_pi(row_idx, rhc2_indices)
        v1 = numeric_first_from_pi(row_idx, rhc1_indices)
        if pd.notna(v2) and pd.notna(v1):
            pi_filtered.iloc[row_idx, delta_idx] = v2 - v1

# ── 7. Summary ────────────────────────────────────────────────────────────────
total_filled  = sum(filled_counts.values())
total_skipped = sum(skipped_counts.values())
print(f"Cells filled from master: {total_filled}")
print(f"Cells skipped (PI value preserved): {total_skipped}")

# Patients in PI sheet with no match in master
unmatched = pi_df[~pi_df[pi_patient_col].isin(patients_in_master)][pi_patient_col].tolist()
if unmatched:
    print(f"WARNING — {len(unmatched)} patients in PI sheet had no match in master: {unmatched}")

# ── 8. Encrypt & Save ─────────────────────────────────────────────────────────
unencrypted_out = io.BytesIO()
with pd.ExcelWriter(unencrypted_out, engine="openpyxl") as writer:
    pi_filtered.to_excel(writer, sheet_name="Sheet1", index=False)

unencrypted_out.seek(0)
encrypted_out = io.BytesIO()
file_to_encrypt = msoffcrypto.OfficeFile(unencrypted_out)
file_to_encrypt.load_key(password=password)
file_to_encrypt.encrypt(password, encrypted_out)

with open(output_path, "wb") as f:
    f.write(encrypted_out.getvalue())

del password
print(f"Done. Saved to: {output_path}")
