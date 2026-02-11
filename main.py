import pandas as pd
import re

print('Loading Datasets...')
print()

df_1 = pd.read_excel("excel/company_dataset_1.xlsx")
df_2 = pd.read_excel("excel/company_dataset_2.xlsx")

print(f"Dataset 1: {len(df_1)} lines")
print(f"Dataset 2: {len(df_2)} lines")
print()

print("First adjusting Dataset 1 column names...")
print()

df_1 = df_1.rename(columns={
    "sStreet1": "address1",
    "sStreet2": "address2",
    "sCity": "city",
    "sProvState": "state",
    "sCountry": "country",
    "sPostalZip": "zip"
})
df_1["address3"] = ""
df_1["ccode"] = ""

print(f"Dataset 1 columns list: {list(df_1.columns)}")
print()

print(" Then dropping duplicates in Dataset 2... ")
print()

before = len(df_2)
df_2 = df_2.drop_duplicates(subset=["custname", "address1", "city", "state", "zip"], keep="first").reset_index(drop=True)

print(f"Dropped {before - len(df_2)} duplicates in DS2. Left: {len(df_2)}")
print()

print(f"Cleaning Company names...")
print()

LEGAL_WORDS = r"\b(inc|ltd|llc|corp|corporation|company|co|limited|partnership)\b"

def clean_company_name(name):
    if pd.isna(name) or str(name).strip() == "":
        return None

    s = str(name).strip().lower()

    m = re.match(r"^(\d{6,})\s+(.*)$", s)
    if m:
        num = m.group(1)
        rest = m.group(2)
        rest = re.sub(r"\(.*?\)", " ", rest)
        rest = re.sub(r"[^\w\s]", " ", rest)
        rest = re.sub(LEGAL_WORDS, " ", rest)
        rest = re.sub(r"\s+", " ", rest).strip()
        key = f"{num} {rest}".strip()
        return key if key else None

    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[0-9]+", " ", s)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(LEGAL_WORDS, " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s if s else None

df_1["custname_clean"] = df_1["custname"].apply(clean_company_name)
df_2["custname_clean"] = df_2["custname"].apply(clean_company_name)


print("Normilizing sate and country...")
print()

CANADA_FULL_TO_CODE = {
    'Ontario': 'ON','Quebec': 'QC','Québec': 'QC','British Columbia': 'BC','Alberta': 'AB','Manitoba': 'MB',
    'Saskatchewan': 'SK','Nova Scotia': 'NS','New Brunswick': 'NB','Newfoundland and Labrador': 'NL',
    'Prince Edward Island': 'PE','Northwest Territories': 'NT','Nunavut': 'NU','Yukon': 'YT'
}
US_FULL_TO_CODE = {
    'Alabama': 'AL','Alaska': 'AK','Arizona': 'AZ','Arkansas': 'AR','California': 'CA','Colorado': 'CO',
    'Connecticut': 'CT','Delaware': 'DE','Florida': 'FL','Georgia': 'GA','Hawaii': 'HI','Idaho': 'ID',
    'Illinois': 'IL','Indiana': 'IN','Iowa': 'IA','Kansas': 'KS','Kentucky': 'KY','Louisiana': 'LA','Maine': 'ME',
    'Maryland': 'MD','Massachusetts': 'MA','Michigan': 'MI','Minnesota': 'MN','Mississippi': 'MS','Missouri': 'MO',
    'Montana': 'MT','Nebraska': 'NE','Nevada': 'NV','New Hampshire': 'NH','New Jersey': 'NJ','New Mexico': 'NM',
    'New York': 'NY','North Carolina': 'NC','North Dakota': 'ND','Ohio': 'OH','Oklahoma': 'OK','Oregon': 'OR',
    'Pennsylvania': 'PA','Rhode Island': 'RI','South Carolina': 'SC','South Dakota': 'SD','Tennessee': 'TN','Texas': 'TX',
    'Utah': 'UT','Vermont': 'VT','Virginia': 'VA','Washington': 'WA','West Virginia': 'WV','Wisconsin': 'WI','Wyoming': 'WY',
    'District of Columbia': 'DC'
}

CANADA_CODES = set(CANADA_FULL_TO_CODE.values())
US_CODES = set(US_FULL_TO_CODE.values())

def normalize_state(x):
    if pd.isna(x) or str(x).strip() == "":
        return None

    s = str(x).strip()

    if "," in s:
        s = s.split(",")[-1].strip()

    if len(s) == 2:
        return s.upper()

    s2 = s.title()
    if s2 in CANADA_FULL_TO_CODE:
        return CANADA_FULL_TO_CODE[s2]
    if s2 in US_FULL_TO_CODE:
        return US_FULL_TO_CODE[s2]
    
    return s

def is_canadian_zip(z):
    if pd.isna(z): return False
    z = str(z).strip().upper().replace(" ", "")
    return bool(re.match(r'^[A-Z]\d[A-Z]\d[A-Z]\d$', z))

def is_us_zip(z):
    if pd.isna(z): return False
    z = str(z).strip()
    return bool(re.match(r'^\d{5}(-\d{4})?$', z))

def normalize_country(c):
    if pd.isna(c) or str(c).strip() == "":
        return None
    s = str(c).strip().lower()
    if s in ["ca", "can", "canada"]:
        return "Canada"
    if s in ["us", "usa", "united states", "united states of america"]:
        return "USA"
    return str(c).strip()

def fill_country(row):
    c = normalize_country(row["country"])
    if c:
        return c

    z = row["zip"]
    st = row["state"]

    if is_canadian_zip(z): return "Canada"
    if is_us_zip(z): return "USA"

    if pd.notna(st):
        st = str(st).upper()
        if st in CANADA_CODES: return "Canada"
        if st in US_CODES: return "USA"

    return None

for df in (df_1, df_2):
    df["state"] = df["state"].apply(normalize_state)
    df["country"] = df.apply(fill_country, axis=1)


print("Normalizing location...")
print()

def make_location(city, state, country):
    city = str(city).strip().lower() if pd.notna(city) else ""
    state = str(state).strip().lower() if pd.notna(state) else ""
    country = str(country).strip().lower() if pd.notna(country) else ""

    if city in ["", "nan"]: city = ""
    if state in ["", "nan"]: state = ""
    if country in ["", "nan"]: country = ""
    if city == "" and state == "":
        return ""

    parts = []
    if city: parts.append(city)
    if state: parts.append(state)
    if country: parts.append(country)
    return " | ".join(parts)

def unique_list(values):
    out = []
    for v in values:
        if isinstance(v, str) and v.strip():
            if v not in out:
                out.append(v)
    return out


print("Creating location per company...")
print()

ds1_locs = (
    df_1.assign(location=df_1.apply(lambda r: make_location(r["city"], r["state"], r["country"]), axis=1))
       .groupby("custname_clean")["location"]
       .apply(unique_list)
       .reset_index()
       .rename(columns={"location": "locations_ds1"})
)

ds2_locs = (
    df_2.assign(location=df_2.apply(lambda r: make_location(r["city"], r["state"], r["country"]), axis=1))
       .groupby("custname_clean")["location"]
       .apply(unique_list)
       .reset_index()
       .rename(columns={"location": "locations_ds2"})
)


print("Creating unique cmpanies from Dataset 1...")
print()

base = df_1[["custname_clean"]].dropna().drop_duplicates().reset_index(drop=True)


print("Creating one-to-many count...")
print()

ds2_match_counts = df_2[df_2["custname_clean"].notna()].groupby("custname_clean").size().reset_index(name="matched_count_ds2")


print("Creating final aggregation table...")
print()

final_df = base.merge(ds1_locs, on="custname_clean", how="left").merge(ds2_locs, on="custname_clean", how="left").merge(ds2_match_counts, on="custname_clean", how="left")

final_df["matched_count_ds2"] = final_df["matched_count_ds2"].fillna(0).astype(int)
final_df["locations_ds1"] = final_df["locations_ds1"].apply(lambda x: x if isinstance(x, list) else [])
final_df["locations_ds2"] = final_df["locations_ds2"].apply(lambda x: x if isinstance(x, list) else [])

final_df["overlapping_locations"] = final_df.apply(lambda r: [x for x in r["locations_ds1"] if x in r["locations_ds2"]], axis=1)
final_df["overlap_count"] = final_df["overlapping_locations"].apply(len)

final_df["locations_ds1"] = final_df["locations_ds1"].apply(lambda x: " || ".join(x))
final_df["locations_ds2"] = final_df["locations_ds2"].apply(lambda x: " || ".join(x))
final_df["overlapping_locations"] = final_df["overlapping_locations"].apply(lambda x: " || ".join(x))


print("Creating metrics...")
print()

total_companies = len(final_df)
matched = (final_df["matched_count_ds2"] > 0).sum()
unmatched = total_companies - matched

match_rate = matched / total_companies * 100 if total_companies else 0
unmatched_rate = unmatched / total_companies * 100 if total_companies else 0

one_to_many = (final_df["matched_count_ds2"] > 1).sum()
one_to_many_rate = one_to_many / matched * 100 if matched else 0

overlap_companies = (final_df["overlap_count"] > 0).sum()
overlap_rate = overlap_companies / matched * 100 if matched else 0


print("STATS")
print()

print(f"Unique companies in DS1: {total_companies}")
print(f"Matched in DS2: {matched} ({match_rate:.1f}%)")
print(f"Unmatched: {unmatched} ({unmatched_rate:.1f}%)")
print(f"One-to-many: {one_to_many} ({one_to_many_rate:.1f}% of matched)")
print(f"Overlap companies: {overlap_companies} ({overlap_rate:.1f}% of matched)")
print()

print("Saving output file...")
print()

final_df = final_df.sort_values(by=["overlap_count", "matched_count_ds2", "custname_clean"], ascending=[False, False, True]).reset_index(drop=True)

output_path = "output_file.csv"
final_df.to_csv(output_path, index=False)

print(f"Saved: {output_path} | rows: {len(final_df)}")
