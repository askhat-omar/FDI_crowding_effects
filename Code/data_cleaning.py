import logging
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pypdf").setLevel(logging.ERROR)

import pandas as pd
from pathlib import Path
from functools import reduce
import camelot

pd.set_option('future.no_silent_downcasting', True)

NA_VALS = ['x', 'х', '-', '', ' ', '...', '…']
INV_COLS = ['rayon', 'l_all_inv', 'f_all_inv']
NON_COLS = ['rayon', 'all_nonextr_inv', 'f_nonextr_inv']


def clean_col(series):
    return (series.astype(str)
            .str.replace(',', '.', regex=False)
            .str.replace(')', '', regex=False)
            .str.replace(' ', '', regex=False)
            .str.replace('\xa0', '', regex=False)
            .str.replace('*', '', regex=False)
            .replace(['nan', 'None', '<NA>'], float('nan'))
            .infer_objects(copy=False)
            .astype(float))


def ctrl_sheet_to_long(path, sheet, header, usecols, nrows, years, rows, rayon_name):
    df = pd.read_excel(path, sheet_name=sheet, header=header, usecols=usecols,
                       nrows=nrows, na_values=NA_VALS)
    df.columns = ['Indicator'] + list(years)
    df = df.loc[list(rows)].reset_index(drop=True)
    df['Indicator'] = ['production', 'agriculture', 'retail']
    for yr in years:
        df[yr] = clean_col(df[yr])
    df = df.melt(id_vars='Indicator', var_name='year', value_name='value')
    df = df.pivot(index='year', columns='Indicator', values='value').reset_index()
    df.insert(0, 'rayon', rayon_name)
    df.columns.name = None
    return df


def broadcast_ppi(ppi_raw, year_list, rayon_list, row=-1):
    vals = ppi_raw.iloc[row].values if isinstance(row, int) else ppi_raw.values[row]
    records = [{'rayon': r, 'year': yr, 'ppi': v}
               for r in rayon_list for yr, v in zip(year_list, vals)]
    df = pd.DataFrame(records)
    df['ppi'] = clean_col(df['ppi'])
    return df


def outer_merge(dfs):
    return reduce(lambda l, r: pd.merge(l, r, on=['rayon', 'year'], how='outer'), dfs)


def to_numeric_inv(df, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(' ', '', regex=False), errors='coerce').fillna(0)
    return df


def process_east_kazakhstan():
    base = Path("../Data/Investments")

    # --- Abay investments 2022-2024 ---
    abay_all_2022 = pd.read_excel(base / "Abay/2022_Investments.xlsx", sheet_name='13', header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    abay_all_2023 = pd.read_excel(base / "Abay/2023_Investments.xlsx", sheet_name='19', header=3, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    abay_all_2024 = pd.read_excel(base / "Abay/2024_Investments.xlsx", sheet_name='14', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    abay_non_2022 = pd.read_excel(base / "Abay/2022_Investments.xlsx", sheet_name='14', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    abay_non_2023 = pd.read_excel(base / "Abay/2023_Investments.xlsx", sheet_name='20', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    abay_non_2024 = pd.read_excel(base / "Abay/2024_Investments.xlsx", sheet_name='15', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    abay_rayon_names = ['Semey_city', 'Kurchatov_city', 'Abay', 'Aksuat', 'Ayagoz', 'Besqaraghay', 'Boroduliha', 'Zharma', 'Kokpekti', 'Urzhar']

    for df in [abay_all_2022, abay_all_2023, abay_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
    for df in [abay_non_2022, abay_non_2023, abay_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)

    # Abay 2024: Zhanasemey->Semey, Makanshi->Urzhar
    for df in [abay_all_2024, abay_non_2024]:
        df.loc[0, df.columns[1]] = df.iloc[0][df.columns[1]] + df.iloc[7][df.columns[1]]
        df.loc[0, df.columns[2]] = df.iloc[0][df.columns[2]] + df.iloc[7][df.columns[2]]
        df.loc[11, df.columns[1]] = df.iloc[10][df.columns[1]] + df.iloc[11][df.columns[1]]
        df.loc[11, df.columns[2]] = df.iloc[10][df.columns[2]] + df.iloc[11][df.columns[2]]
        df.drop([7, 10], inplace=True)
        df.reset_index(drop=True, inplace=True)
        df['rayon'] = abay_rayon_names

    for df in [abay_all_2022, abay_all_2023]:
        df['rayon'] = abay_rayon_names
    for df in [abay_non_2022, abay_non_2023]:
        df['rayon'] = abay_rayon_names

    abay_2022 = abay_all_2022.merge(abay_non_2022, on='rayon', how='left')
    abay_2023 = abay_all_2023.merge(abay_non_2023, on='rayon', how='left')
    abay_2024 = abay_all_2024.merge(abay_non_2024, on='rayon', how='left')

    # --- EK investments 2022-2024 ---
    ek_all_2022 = pd.read_excel(base / "East-Kazakhstan/2022_Investments.xlsx", sheet_name='13', header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    ek_all_2023 = pd.read_excel(base / "East-Kazakhstan/2023_Investments.xlsx", sheet_name='19', header=4, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    ek_all_2024 = pd.read_excel(base / "East-Kazakhstan/2024_Investments.xlsx", sheet_name='14', header=4, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    ek_non_2022 = pd.read_excel(base / "East-Kazakhstan/2022_Investments.xlsx", sheet_name='14', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    ek_non_2023 = pd.read_excel(base / "East-Kazakhstan/2023_Investments.xlsx", sheet_name='20', header=4, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    ek_non_2024 = pd.read_excel(base / "East-Kazakhstan/2024_Investments.xlsx", sheet_name='15', header=4, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    ek_rayon_names = ['Oskemen_city', 'Ridder_city', 'Glubokoe', 'Zaysan', 'Altay', 'Kurshim', 'Qatonqaraghay', 'Samar', 'Tarbagatay', 'Ulan', 'Shemonaiha']

    for df in [ek_all_2022, ek_all_2023, ek_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
    for df in [ek_non_2022, ek_non_2023, ek_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)

    # EK 2024: Markakol->Kurshim, Ulken Naryn->Katonkaragay
    for df in [ek_all_2024, ek_non_2024]:
        df.loc[5, df.columns[1]] = df.iloc[5][df.columns[1]] + df.iloc[7][df.columns[1]]
        df.loc[5, df.columns[2]] = df.iloc[5][df.columns[2]] + df.iloc[7][df.columns[2]]
        df.loc[6, df.columns[1]] = df.iloc[6][df.columns[1]] + df.iloc[11][df.columns[1]]
        df.loc[6, df.columns[2]] = df.iloc[6][df.columns[2]] + df.iloc[11][df.columns[2]]
        df.drop([7, 11], inplace=True)
        df.reset_index(drop=True, inplace=True)
        df['rayon'] = ek_rayon_names

    for df in [ek_all_2022, ek_all_2023]:
        df['rayon'] = ek_rayon_names
    for df in [ek_non_2022, ek_non_2023]:
        df['rayon'] = ek_rayon_names

    ek_2022 = ek_all_2022.merge(ek_non_2022, on='rayon', how='left')
    ek_2023 = ek_all_2023.merge(ek_non_2023, on='rayon', how='left')
    ek_2024 = ek_all_2024.merge(ek_non_2024, on='rayon', how='left')

    def concat_abay_ek(abay_df, ek_df):
        df = pd.concat([abay_df, ek_df], ignore_index=True)
        # Aksuat->Tarbagatay (abay loc 3 -> ek Tarbagatay which is loc 18 in combined)
        df.iloc[18, 1] = df.iloc[18][df.columns[1]] + df.iloc[3][df.columns[1]]
        df.iloc[18, 2] = df.iloc[18][df.columns[2]] + df.iloc[3][df.columns[2]]
        # Samar->Kokpekti (ek Samar is loc 17 -> abay Kokpekti is loc 8)
        df.iloc[8, 1] = df.iloc[8][df.columns[1]] + df.iloc[17][df.columns[1]]
        df.iloc[8, 2] = df.iloc[8][df.columns[2]] + df.iloc[17][df.columns[2]]
        df.drop([3, 17], inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df

    east_2022 = concat_abay_ek(abay_2022, ek_2022)
    east_2023 = concat_abay_ek(abay_2023, ek_2023)
    east_2024 = concat_abay_ek(abay_2024, ek_2024)

    # --- EK 2018 (camelot PDF) ---
    east_rayon_names_19 = ['Oskemen_city', 'Kurchatov_city', 'Ridder_city', 'Semey_city', 'Abay', 'Ayagoz', 'Besqaraghay', 'Boroduliha', 'Glubokoe', 'Zharma', 'Zaysan', 'Altay', 'Kokpekti', 'Kurshim', 'Qatonqaraghay', 'Tarbagatay', 'Ulan', 'Urzhar', 'Shemonaiha']

    _pdf = Path("../Data/Investments/East-Kazakhstan/2018_Investments.pdf")
    east_all_2018 = camelot.read_pdf(_pdf, pages='28', flavor='stream')[0].df
    east_nonextr_2018 = camelot.read_pdf(_pdf, pages='29', flavor='stream')[0].df

    east_all_2018 = east_all_2018[[0, 2, 5]].drop([0, 1, 2, 3, 4]).reset_index(drop=True)
    east_all_2018.columns = INV_COLS
    to_numeric_inv(east_all_2018, ['l_all_inv', 'f_all_inv'])
    east_all_2018['l_all_inv'] = east_all_2018['l_all_inv'].astype(str).str.replace(' ', '', regex=False).astype(int)
    east_all_2018['f_all_inv'] = east_all_2018['f_all_inv'].astype(str).str.replace(' ', '', regex=False).astype(int)
    east_all_2018['rayon'] = east_rayon_names_19

    east_nonextr_2018 = east_nonextr_2018[[0, 2, 5]].drop([0, 1, 2, 3, 4]).reset_index(drop=True)
    east_nonextr_2018.columns = NON_COLS
    east_nonextr_2018['all_nonextr_inv'] = east_nonextr_2018['all_nonextr_inv'].astype(str).str.replace(' ', '', regex=False).astype(int)
    east_nonextr_2018['f_nonextr_inv'] = east_nonextr_2018['f_nonextr_inv'].astype(str).str.replace(' ', '', regex=False).astype(int)
    to_numeric_inv(east_nonextr_2018, ['all_nonextr_inv', 'f_nonextr_inv'])
    east_nonextr_2018['l_nonextr_inv'] = east_nonextr_2018['all_nonextr_inv'] - east_nonextr_2018['f_nonextr_inv']
    east_nonextr_2018 = east_nonextr_2018.drop('all_nonextr_inv', axis=1)
    east_nonextr_2018['rayon'] = east_rayon_names_19

    # --- EK 2019-2021 (Excel) ---
    east_all_2019 = pd.read_excel(base / "East-Kazakhstan/2019_Investments.xls", sheet_name='14', header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    east_all_2020 = pd.read_excel(base / "East-Kazakhstan/2020_Investments.xlsx", sheet_name='13', header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    east_all_2021 = pd.read_excel(base / "East-Kazakhstan/2021_Investments.xlsx", sheet_name='13', header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    east_non_2019 = pd.read_excel(base / "East-Kazakhstan/2019_Investments.xls", sheet_name='15', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop([0, 20, 21, 22]).reset_index(drop=True)
    east_non_2020 = pd.read_excel(base / "East-Kazakhstan/2020_Investments.xlsx", sheet_name='14', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    east_non_2021 = pd.read_excel(base / "East-Kazakhstan/2021_Investments.xlsx", sheet_name='14', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [east_all_2019, east_all_2020, east_all_2021]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = east_rayon_names_19
    for df in [east_non_2019, east_non_2020, east_non_2021]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = east_rayon_names_19

    east_2018 = east_all_2018.merge(east_nonextr_2018, on='rayon')
    east_2019 = east_all_2019.merge(east_non_2019, on='rayon', how='left')
    east_2020 = east_all_2020.merge(east_non_2020, on='rayon', how='left')
    east_2021 = east_all_2021.merge(east_non_2021, on='rayon', how='left')

    dfs_inv = [east_2018, east_2019, east_2020, east_2021, east_2022, east_2023, east_2024]
    for i, df in enumerate(dfs_inv):
        df['oblast'] = 'East-Kazakhstan'
        df['year'] = 2018 + i

    east_inv = pd.concat(dfs_inv, ignore_index=True)
    east_inv['rayon'] = east_inv['rayon'].str.strip()

    # --- Employment ---
    # 2022-2024: combine Abay + EK proper (Samar->Kokpekti, Aksuat->Tarbagatay)
    abay_empl = pd.read_excel(Path("../Data/Employment/Abay.xlsx"), skiprows=22, nrows=10, header=None, usecols=[0, 1, 2, 3], na_values=NA_VALS).fillna(0).reset_index(drop=True)
    abay_empl.columns = ['rayon', 2022, 2023, 2024]
    abay_empl['rayon'] = abay_rayon_names
    for yr in [2022, 2023, 2024]:
        abay_empl[yr] = clean_col(abay_empl[yr])

    ek_empl_22_24 = pd.read_excel(Path("../Data/Employment/East-Kazakhstan.xlsx"), skiprows=30, nrows=20, header=None, usecols='A, J, K, L', na_values=NA_VALS).dropna().reset_index(drop=True)
    ek_empl_22_24.columns = ['rayon', 2022, 2023, 2024]
    ek_empl_22_24['rayon'] = ['Oskemen_city', 'Ridder_city', 'Glubokoe', 'Zaysan', 'Altay', 'Samar', 'Kurshim', 'Qatonqaraghay', 'Tarbagatay', 'Ulan', 'Shemonaiha']
    for yr in [2022, 2023, 2024]:
        ek_empl_22_24[yr] = ek_empl_22_24[yr].astype(float)

    # Abay idx 8=Kokpekti, EK idx 5(+10offset)=15=Samar; Abay idx 3=Aksuat, EK idx 8(+10offset)=18=Tarbagatay
    east_empl_22_24 = pd.concat([abay_empl, ek_empl_22_24], ignore_index=True)
    east_empl_22_24.loc[8, [2022, 2023, 2024]] += east_empl_22_24.loc[15, [2022, 2023, 2024]].values
    east_empl_22_24.loc[18, [2022, 2023, 2024]] += east_empl_22_24.loc[3, [2022, 2023, 2024]].values
    east_empl_22_24 = east_empl_22_24.drop([3, 15]).reset_index(drop=True)
    east_empl_22_24_long = east_empl_22_24.melt(id_vars='rayon', var_name='year', value_name='employment')

    # 2018-2021: full pre-split EK (19 rayons)
    ek_empl_18_21 = pd.read_excel(Path("../Data/Employment/East-Kazakhstan.xlsx"), skiprows=30, nrows=20, header=None, usecols='A, F:I', na_values=NA_VALS).dropna().reset_index(drop=True)
    ek_empl_18_21.columns = ['rayon', 2018, 2019, 2020, 2021]
    ek_empl_18_21['rayon'] = ['Oskemen_city', 'Kurchatov_city', 'Ridder_city', 'Semey_city', 'Abay', 'Ayagoz', 'Besqaraghay', 'Boroduliha', 'Glubokoe', 'Zharma', 'Zaysan', 'Altay', 'Kokpekti', 'Kurshim', 'Qatonqaraghay', 'Tarbagatay', 'Ulan', 'Urzhar', 'Shemonaiha']
    for yr in [2018, 2019, 2020, 2021]:
        ek_empl_18_21[yr] = ek_empl_18_21[yr].astype(float)
    east_empl_18_21_long = ek_empl_18_21.melt(id_vars='rayon', var_name='year', value_name='employment')

    east_employment = pd.concat([east_empl_18_21_long, east_empl_22_24_long], ignore_index=True)

    # --- Controls ---
    ctrl_abay_path = Path("../Data/Controls/Abay_All.xls")
    ctrl_ek_path = Path("../Data/Controls/East-Kazakhstan_All.xls")
    abay_yrs = [2018, 2019, 2020, 2021, 2022, 2023, 2024]
    ek_yrs = [2018, 2019, 2020, 2021, 2022, 2023]

    def _wide(path, sheet, header, usecols, nrows, years, rows):
        """Read a control sheet and return 3-row wide DataFrame (production, agriculture, retail) x years."""
        df = pd.read_excel(path, sheet_name=sheet, header=header, usecols=usecols, nrows=nrows, na_values=NA_VALS)
        df.columns = ['Indicator'] + list(years)
        df = df.loc[list(rows)].reset_index(drop=True)
        df['Indicator'] = ['production', 'agriculture', 'retail']
        for yr in years:
            df[yr] = clean_col(df[yr])
        return df

    def _wide_to_long(df, rayon_name):
        """Convert wide 3-row ctrl df to long format with rayon column."""
        out = pd.melt(df, id_vars=['Indicator'], var_name='year', value_name='value')
        out = out.pivot(index='year', columns='Indicator', values='value').reset_index()
        out.insert(0, 'rayon', rayon_name)
        out.columns.name = None
        return out

    def _abay(sheet, header, usecols, nrows, rows):
        return _wide(ctrl_abay_path, sheet, header, usecols, nrows, abay_yrs, rows)

    def _ek(sheet, rows):
        return _wide(ctrl_ek_path, sheet, 4, 'A, AC:AH', 238, ek_yrs, rows)

    # --- Abay rayons (2018-2024) ---
    kurchatov = _wide_to_long(_abay('моног.Курчатов', 2, 'B, AD:AJ', 272, [134, 167, 257]), 'Kurchatov_city')

    semey_w = _abay('Семей', 2, 'A, AC:AI', 229, [117, 149, 218])
    zh_semey_w = _wide(ctrl_abay_path, 'Жаңасемей', 4, 'A:B', 229, [2024], [117, 149, 218])
    semey_w[2024] = semey_w[2024] + zh_semey_w[2024]
    semey = _wide_to_long(semey_w, 'Semey_city')

    abay_ctrl = _wide_to_long(_abay('Абай', 3, 'A, AC:AI', 230, [117, 149, 218]), 'Abay')
    ayagoz = _wide_to_long(_abay('Аягоз', 2, 'A, AC:AI', 230, [117, 149, 218]), 'Ayagoz')
    beskaragay = _wide_to_long(_abay('Бескарагай', 4, 'A, AC:AI', 232, [119, 151, 220]), 'Besqaraghay')
    boroduliha = _wide_to_long(_abay('Бородулиха', 4, 'A, AC:AI', 230, [117, 149, 218]), 'Boroduliha')
    zharma = _wide_to_long(_abay('Жарма', 4, 'A, AC:AI', 230, [117, 149, 218]), 'Zharma')

    kokpekti_w = _abay('Кокпекты', 4, 'A, AC:AI', 232, [119, 151, 220])
    aksuat_w = _abay('Аксуат', 4, 'A, AC:AI', 230, [117, 149, 218])

    urzhar_w = _abay('Урджар', 4, 'A, AC:AI', 230, [117, 149, 218])
    makanshi_w = _wide(ctrl_abay_path, 'Мақаншы', 4, 'A:B', 230, [2024], [117, 149, 218])
    urzhar_w[2024] = urzhar_w[2024] + makanshi_w[2024]
    urzhar_ctrl = _wide_to_long(urzhar_w, 'Urzhar')

    # --- EK rayons (2018-2023) ---
    oskemen = _wide_to_long(_ek('У-ка', [125, 157, 226]), 'Oskemen_city')

    ridder_w = _ek('Риддер', [125, 157, 226])
    ridder_mono_w = _wide(ctrl_ek_path, 'Моногорода Риддер', 1, 'B, K:P', 279, ek_yrs, [137, 172, 263])
    ridder_mono_w[ek_yrs] = ridder_mono_w[ek_yrs] + ridder_w[ek_yrs]
    ridder_city_ctrl = _wide_to_long(ridder_mono_w, 'Ridder_city')

    glubokoe = _wide_to_long(_ek('Глубокое', [125, 157, 226]), 'Glubokoe')
    zaysan = _wide_to_long(_ek('Зайсан', [125, 157, 226]), 'Zaysan')

    altay_w = _ek('Алтай', [125, 157, 226])
    altay_mono_w = _wide(ctrl_ek_path, 'Моногорода Алтай', 1, 'B, K:P', 278, ek_yrs, [137, 172, 263])
    altay_w[ek_yrs] = altay_w[ek_yrs] + altay_mono_w[ek_yrs].fillna(0)
    altay_ctrl = _wide_to_long(altay_w, 'Altay')

    samar_w = _ek('Самар', [125, 157, 226])
    kokpekti_w[ek_yrs] = kokpekti_w[ek_yrs] + samar_w[ek_yrs].fillna(0)
    kokpekti_ctrl = _wide_to_long(kokpekti_w, 'Kokpekti')

    kurshim = _wide_to_long(_ek('Курчум', [125, 157, 226]), 'Kurshim')
    katon = _wide_to_long(_ek('Катон-Карагай', [125, 157, 226]), 'Qatonqaraghay')

    tarbagatay_w = _ek('Тарбагатай', [125, 157, 226])
    tarbagatay_w[ek_yrs] = tarbagatay_w[ek_yrs] + aksuat_w[ek_yrs].fillna(0)
    tarbagatay_ctrl = _wide_to_long(tarbagatay_w, 'Tarbagatay')

    ulan = _wide_to_long(_ek('Улан', [125, 157, 226]), 'Ulan')
    shemonaiha = _wide_to_long(_ek('Шемонаиха', [125, 157, 226]), 'Shemonaiha')

    east_controls = pd.concat([oskemen, kurchatov, ridder_city_ctrl, semey, abay_ctrl,
                                ayagoz, beskaragay, boroduliha, glubokoe, zharma, zaysan,
                                altay_ctrl, kokpekti_ctrl, kurshim, katon, tarbagatay_ctrl,
                                ulan, urzhar_ctrl, shemonaiha], ignore_index=True)

    # --- PPI ---
    east_ppi_raw = pd.read_excel(Path("../Data/PPI/East-Kazakhstan.xls"), header=4, usecols='AC:AH', nrows=238, na_values=NA_VALS)
    east_ppi_raw.columns = list(range(2018, 2024))
    ppi_vals = east_ppi_raw.iloc[80].values.copy()
    # Hardcoded corrections for rows where camelot misread values
    corrected = ['100.7', '103.4', '121.1', '116.6', float('nan'), '101.5']
    for i, v in enumerate(corrected):
        if v is not float('nan') and v != float('nan'):
            ppi_vals[i] = v

    rayon_list = list(east_inv['rayon'].unique())
    year_list = list(range(2018, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    east_ppi = pd.DataFrame(ppi_records)
    east_ppi['ppi'] = clean_col(east_ppi['ppi'])

    east_merged = outer_merge([east_inv, east_employment, east_controls, east_ppi])
    east_merged['oblast'] = east_merged['oblast'].fillna('East-Kazakhstan')
    return east_merged


def process_almaty():
    base = Path("../Data/Investments")

    # --- Zhetisu investments 2022-2024 ---
    zhetisu_all_2022 = pd.read_excel(base / "Zhetisu/2022_Investments.xls", sheet_name='13.', header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    zhetisu_all_2023 = pd.read_excel(base / "Zhetisu/2023_Investments.xlsx", sheet_name='19', header=3, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    zhetisu_all_2024 = pd.read_excel(base / "Zhetisu/2024_Investments.xls", sheet_name='14', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    zhetisu_non_2022 = pd.read_excel(base / "Zhetisu/2022_Investments.xls", sheet_name='14.', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    zhetisu_non_2023 = pd.read_excel(base / "Zhetisu/2023_Investments.xlsx", sheet_name='20', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    zhetisu_non_2024 = pd.read_excel(base / "Zhetisu/2024_Investments.xls", sheet_name='15', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    zhetisu_rayon_names = ['Taldyqorgan_city', 'Tekeli_city', 'Aqsu', 'Alakol', 'Yeskeldi', 'Kerbulaq', 'Koksu', 'Qaratal', 'Panfilov', 'Sarqan']

    for df in [zhetisu_all_2022, zhetisu_all_2023, zhetisu_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = zhetisu_rayon_names
    for df in [zhetisu_non_2022, zhetisu_non_2023, zhetisu_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = zhetisu_rayon_names

    zhetisu_2022 = zhetisu_all_2022.merge(zhetisu_non_2022, on='rayon', how='left')
    zhetisu_2023 = zhetisu_all_2023.merge(zhetisu_non_2023, on='rayon', how='left')
    zhetisu_2024 = zhetisu_all_2024.merge(zhetisu_non_2024, on='rayon', how='left')

    # --- Almaty investments 2022-2024 ---
    alm_rayon_names = ['Qonayev_city', 'Balqash', 'Enbekshiqazaq', 'Zhambyl', 'Kegen', 'Qarasay', 'Rayimbek', 'Talgar', 'Uyghur', 'Ile']
    alm_rayon_names_2024 = ['Qonayev_city', 'Alatau_city', 'Balqash', 'Enbekshiqazaq', 'Zhambyl', 'Kegen', 'Qarasay', 'Rayimbek', 'Talgar', 'Uyghur', 'Ile']

    alm_all_2022 = pd.read_excel(base / "Almaty/2022_Investments.xls", sheet_name='13', header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2023 = pd.read_excel(base / "Almaty/2023_Investments.xlsx", sheet_name='14', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2024 = pd.read_excel(base / "Almaty/2024_Investments.xls", sheet_name='14', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    alm_non_2022 = pd.read_excel(base / "Almaty/2022_Investments.xls", sheet_name='14.', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:9]
    alm_non_2023 = pd.read_excel(base / "Almaty/2023_Investments.xlsx", sheet_name='15', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_non_2024 = pd.read_excel(base / "Almaty/2024_Investments.xls", sheet_name='15', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [alm_all_2022, alm_all_2023]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = alm_rayon_names
    alm_all_2024.columns = INV_COLS
    to_numeric_inv(alm_all_2024, ['l_all_inv', 'f_all_inv'])
    alm_all_2024['rayon'] = alm_rayon_names_2024

    for df in [alm_non_2022, alm_non_2023, alm_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)

    for df in [alm_non_2022, alm_non_2023]:
        df['rayon'] = alm_rayon_names
    alm_non_2024['rayon'] = alm_rayon_names_2024

    # Kegen->Rayimbek: 2022 loc[6] += loc[4], 2023 loc[6] += loc[4], 2024 loc[7] += loc[5]
    def merge_kegen(all_df, non_df, kegen_idx, rayimbek_idx, names):
        for df in [all_df, non_df]:
            for col in df.columns[1:]:
                df.loc[rayimbek_idx, col] = df.iloc[rayimbek_idx][col] + df.iloc[kegen_idx][col]
            df.drop(kegen_idx, inplace=True)
            df.reset_index(drop=True, inplace=True)
        names_filtered = [n for i, n in enumerate(names) if i != kegen_idx]
        all_df['rayon'] = names_filtered
        non_df['rayon'] = names_filtered
        return all_df.merge(non_df, on='rayon', how='left')

    alm_2022 = merge_kegen(alm_all_2022, alm_non_2022, 4, 6, alm_rayon_names)
    alm_2023 = merge_kegen(alm_all_2023, alm_non_2023, 4, 6, alm_rayon_names)
    alm_2024 = merge_kegen(alm_all_2024, alm_non_2024, 5, 7, alm_rayon_names_2024)
    # Alatau_city was split from Ile rayon in 2024 — merge back into Ile
    inv_cols_num = ['l_all_inv', 'f_all_inv', 'f_nonextr_inv', 'l_nonextr_inv']
    ile_mask = alm_2024['rayon'] == 'Ile'
    ala_mask = alm_2024['rayon'] == 'Alatau_city'
    if ala_mask.any() and ile_mask.any():
        alm_2024.loc[ile_mask, inv_cols_num] = (
            alm_2024.loc[ile_mask, inv_cols_num].values +
            alm_2024.loc[ala_mask, inv_cols_num].values)
        alm_2024 = alm_2024[~ala_mask].reset_index(drop=True)

    # --- Almaty 2016-2021 ---
    # Row orders match actual Excel file structure (verified against reference notebook)
    rayons_2016 = ['Aqsu', 'Alakol', 'Balqash', 'Enbekshiqazaq', 'Yeskeldi', 'Zhambyl', 'Ile', 'Qarasay', 'Qaratal', 'Kerbulaq', 'Koksu', 'Panfilov', 'Rayimbek', 'Sarqan', 'Talgar', 'Uyghur', 'Qonayev_city', 'Taldyqorgan_city', 'Tekeli_city']
    rayons_2017 = ['Taldyqorgan_city', 'Qonayev_city', 'Tekeli_city', 'Aqsu', 'Alakol', 'Balqash', 'Enbekshiqazaq', 'Zhambyl', 'Kerbulaq', 'Koksu', 'Qaratal', 'Qarasay', 'Panfilov', 'Rayimbek', 'Sarqan', 'Talgar', 'Yeskeldi', 'Uyghur', 'Ile']
    # Kegen at index 8, Rayimbek at index 14 — matches merge logic below
    rayons_w_kegen = ['Taldyqorgan_city', 'Qonayev_city', 'Tekeli_city', 'Aqsu', 'Alakol', 'Balqash', 'Enbekshiqazaq', 'Zhambyl', 'Kegen', 'Kerbulaq', 'Koksu', 'Qaratal', 'Qarasay', 'Panfilov', 'Rayimbek', 'Sarqan', 'Talgar', 'Yeskeldi', 'Uyghur', 'Ile']

    alm_all_2016 = pd.read_excel(base / "Almaty/2016_Investments/Internal_vs_external.xls", header=6, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2017 = pd.read_excel(base / "Almaty/2017_Investments/Internal_vs_external.xls", header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2018 = pd.read_excel(base / "Almaty/2018_Investments/Internal_vs_external.xls", header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2019 = pd.read_excel(base / "Almaty/2019_Investments.xls", sheet_name='14.', header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2020 = pd.read_excel(base / "Almaty/2020_Investments.xls", sheet_name='13.', header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    alm_all_2021 = pd.read_excel(base / "Almaty/2021_Investments.xls", sheet_name='13.', header=5, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    alm_non_2016 = pd.read_excel(base / "Almaty/2016_Investments/Greenfield_investments.xls", header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:18]
    alm_non_2017 = pd.read_excel(base / "Almaty/2017_Investments/Greenfield_investments.xls", header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:18]
    alm_non_2018 = pd.read_excel(base / "Almaty/2018_Investments/Greenfield_investments.xlsx", header=5, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:19]
    alm_non_2019 = pd.read_excel(base / "Almaty/2019_Investments.xls", sheet_name='15.', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:19]
    alm_non_2020 = pd.read_excel(base / "Almaty/2020_Investments.xls", sheet_name='14.', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).loc[:19]
    alm_non_2021 = pd.read_excel(base / "Almaty/2021_Investments.xls", sheet_name='14.', header=6, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [alm_all_2016, alm_all_2017]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
    for df in [alm_all_2018, alm_all_2019, alm_all_2020, alm_all_2021]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
    for df in [alm_non_2016, alm_non_2017, alm_non_2018, alm_non_2019, alm_non_2020, alm_non_2021]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)

    alm_all_2016['rayon'] = rayons_2016
    alm_all_2017['rayon'] = rayons_2017
    alm_non_2016['rayon'] = rayons_2016
    alm_non_2017['rayon'] = rayons_2017

    # 2018-2021 have Kegen; Kegen->Rayimbek: loc[14] += loc[8], drop 8
    for all_df, non_df in [(alm_all_2018, alm_non_2018), (alm_all_2019, alm_non_2019), (alm_all_2020, alm_non_2020), (alm_all_2021, alm_non_2021)]:
        all_df['rayon'] = rayons_w_kegen
        non_df['rayon'] = rayons_w_kegen
        for df in [all_df, non_df]:
            for col in df.columns[1:]:
                try:
                    df.loc[14, col] = df.iloc[14][col] + df.iloc[8][col]
                except Exception:
                    pass
            df.drop(8, inplace=True)
            df.reset_index(drop=True, inplace=True)
        names_no_kegen = [n for i, n in enumerate(rayons_w_kegen) if i != 8]
        all_df['rayon'] = names_no_kegen
        non_df['rayon'] = names_no_kegen

    alm_early_list = [
        alm_all_2016.merge(alm_non_2016, on='rayon', how='left'),
        alm_all_2017.merge(alm_non_2017, on='rayon', how='left'),
        alm_all_2018.merge(alm_non_2018, on='rayon', how='left'),
        alm_all_2019.merge(alm_non_2019, on='rayon', how='left'),
        alm_all_2020.merge(alm_non_2020, on='rayon', how='left'),
        alm_all_2021.merge(alm_non_2021, on='rayon', how='left'),
    ]
    for i, df in enumerate(alm_early_list):
        df['oblast'] = 'Almaty'
        df['year'] = 2016 + i

    combined_late = pd.concat([zhetisu_2022, alm_2022], ignore_index=True)
    combined_late['oblast'] = 'Almaty'
    combined_late['year'] = 2022
    combined_late2 = pd.concat([zhetisu_2023, alm_2023], ignore_index=True)
    combined_late2['oblast'] = 'Almaty'
    combined_late2['year'] = 2023
    combined_late3 = pd.concat([zhetisu_2024, alm_2024], ignore_index=True)
    combined_late3['oblast'] = 'Almaty'
    combined_late3['year'] = 2024

    almaty_inv = pd.concat(alm_early_list + [combined_late, combined_late2, combined_late3], ignore_index=True)
    almaty_inv['rayon'] = almaty_inv['rayon'].str.strip()

    # --- Employment ---
    zhetisu_empl = pd.read_excel(Path("../Data/Employment/Zhetisu.xlsx"), skiprows=21, nrows=10, header=None, usecols='A:C', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    zhetisu_empl.columns = ['rayon', 2022, 2023]
    zhetisu_empl['rayon'] = zhetisu_rayon_names
    zhetisu_empl_long = zhetisu_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    alm_empl_rayon_map = {
        'Талдықорған қ.': 'Taldyqorgan_city', 'Талдықорған қ.ә.': 'Taldyqorgan_city',
        'г.а. Талдыкорган': 'Taldyqorgan_city', 'г.а. Талдықорған': 'Taldyqorgan_city',
        'Текелі қ.': 'Tekeli_city', 'Текелі қ.ә.': 'Tekeli_city',
        'г.а. Текели': 'Tekeli_city', 'г.а. Текелі': 'Tekeli_city',
        'Қонаев қ.': 'Qonayev_city', 'Қонаев қ.ә.': 'Qonayev_city',
        'г.а. Конаев': 'Qonayev_city', 'г.а. Қонаев': 'Qonayev_city',
        'Балқаш ауданы': 'Balqash', 'Балқашский': 'Balqash', 'Балхашский': 'Balqash',
        'Аксуский': 'Aqsu', 'Аксуский район': 'Aqsu',
        'Алакольский': 'Alakol', 'Алакольский район': 'Alakol',
        'Балхашский': 'Balqash', 'Балхашский район': 'Balqash',
        'Енбекшиказахский': 'Enbekshiqazaq', 'Енбекшіқазақ ауданы': 'Enbekshiqazaq',
        'Жамбылский': 'Zhambyl', 'Жамбыл ауданы': 'Zhambyl',
        'Кегенский': 'Kegen', 'Кеген ауданы': 'Kegen',
        'Кербулакский': 'Kerbulaq', 'Кербұлақ ауданы': 'Kerbulaq',
        'Коксуский': 'Koksu', 'Қоқсу ауданы': 'Koksu',
        'Каратальский': 'Qaratal', 'Қаратал ауданы': 'Qaratal',
        'Карасайский': 'Qarasay', 'Қарасай ауданы': 'Qarasay',
        'Панфиловский': 'Panfilov', 'Панфилов ауданы': 'Panfilov',
        'Райымбекский': 'Rayimbek', 'Райымбек ауданы': 'Rayimbek',
        'Саркандский': 'Sarqan', 'Сарқан ауданы': 'Sarqan',
        'Талгарский': 'Talgar', 'Талғар ауданы': 'Talgar',
        'Ескельдинский': 'Yeskeldi', 'Ескелді ауданы': 'Yeskeldi',
        'Уйгурский': 'Uyghur', 'Ұйғыр ауданы': 'Uyghur',
        'Илийский': 'Ile', 'Іле ауданы': 'Ile',
    }

    alm_empl_22_23 = pd.read_excel(Path("../Data/Employment/Almaty.xlsx"), skiprows=31, nrows=20, header=None, usecols='A, J, K', na_values=NA_VALS).dropna().reset_index(drop=True)
    alm_empl_22_23.columns = ['rayon', 2022, 2023]
    alm_empl_22_23['rayon'] = ['Qonayev_city', 'Balqash', 'Enbekshiqazaq', 'Zhambyl', 'Kegen', 'Qarasay', 'Rayimbek', 'Talgar', 'Uyghur', 'Ile']
    # Kegen->Rayimbek: Rayimbek is at idx 6, Kegen is at idx 4
    alm_empl_22_23.loc[6, 2022] = alm_empl_22_23.iloc[6][2022] + alm_empl_22_23.iloc[4][2022]
    alm_empl_22_23.loc[6, 2023] = alm_empl_22_23.iloc[6][2023] + alm_empl_22_23.iloc[4][2023]
    alm_empl_22_23.drop(4, inplace=True)
    alm_empl_22_23.reset_index(drop=True, inplace=True)
    alm_empl_22_23_long = alm_empl_22_23.melt(id_vars='rayon', var_name='year', value_name='employment')

    alm_empl_16_21 = pd.read_excel(Path("../Data/Employment/Almaty.xlsx"), skiprows=31, nrows=20, header=None, usecols='A, D, E, F, G, H, I', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    alm_empl_16_21.columns = ['rayon', 2016, 2017, 2018, 2019, 2020, 2021]
    alm_empl_16_21['rayon'] = alm_empl_16_21['rayon'].str.strip().replace(alm_empl_rayon_map)
    # Kegen->Rayimbek: loc[14] += loc[8], drop 8
    for yr in [2016, 2017, 2018, 2019, 2020, 2021]:
        alm_empl_16_21.loc[14, yr] = alm_empl_16_21.iloc[14][yr] + alm_empl_16_21.iloc[8][yr]
    alm_empl_16_21.drop(8, inplace=True)
    alm_empl_16_21.reset_index(drop=True, inplace=True)
    alm_empl_16_21_long = alm_empl_16_21.melt(id_vars='rayon', var_name='year', value_name='employment')

    almaty_employment = pd.concat([zhetisu_empl_long, alm_empl_22_23_long, alm_empl_16_21_long], ignore_index=True)

    # --- Controls ---
    ctrl_zhetisu_path = Path("../Data/Controls/Zhetisu_All.xls")
    ctrl_alm_path = Path("../Data/Controls/Almaty_All.xls")

    zhetisu_ctrl_years = [2022, 2023, 2024]
    zhetisu_ctrl_sheets = {
        'Taldyqorgan_city': 'г.а.Талдыкорган',
        'Tekeli_city': 'г.а.Текели',
        'Aqsu': 'Аксуский',
        'Alakol': 'Алакольский',
        'Yeskeldi': 'Ескельдинский',
        'Kerbulaq': 'Кербулакский',
        'Koksu': 'Коксуский',
        'Qaratal': 'Каратальский',
        'Panfilov': 'Панфиловский',
        'Sarqan': 'Сарканский',
    }
    zhetisu_ctrl_dfs = []
    for rayon_name, sheet in zhetisu_ctrl_sheets.items():
        df = ctrl_sheet_to_long(ctrl_zhetisu_path, sheet, 3, 'B, C:E', 130, zhetisu_ctrl_years, [56, 77, 125], rayon_name)
        zhetisu_ctrl_dfs.append(df)

    alm_ctrl_years = list(range(2016, 2024))
    alm_ctrl_sheets = {
        'Qonayev_city': 'г.а.Конаев',
        'Balqash': 'Балхашский',
        'Enbekshiqazaq': 'Енбекшиказахский',
        'Zhambyl': 'Жамбылский',
        'Qarasay': 'Карасайский',
        'Rayimbek': 'Райымбекский',
        'Talgar': 'Талгарский',
        'Uyghur': 'Уйгурский',
        'Ile': 'Илийский',
    }
    alm_ctrl_dfs = []
    for rayon_name, sheet in alm_ctrl_sheets.items():
        df = ctrl_sheet_to_long(ctrl_alm_path, sheet, 3, 'B, AB:AI', 130, alm_ctrl_years, [56, 77, 125], rayon_name)
        alm_ctrl_dfs.append(df)

    # Kegen: usecols='B, C:H', years=[2018..2023]
    kegen_years = list(range(2018, 2024))
    kegen_ctrl = ctrl_sheet_to_long(ctrl_alm_path, 'Кегенский', 3, 'B, C:H', 130, kegen_years, [56, 77, 125], 'Kegen_temp')
    rayimbek_ctrl = next((d for d in alm_ctrl_dfs if d['rayon'].iloc[0] == 'Rayimbek'), None)
    if rayimbek_ctrl is not None:
        for yr in kegen_years:
            rayimbek_ctrl.loc[rayimbek_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += kegen_ctrl.loc[kegen_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values

    almaty_controls = pd.concat(zhetisu_ctrl_dfs + alm_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    alm_ppi_raw = pd.read_excel(Path("../Data/PPI/Almaty.xls"), header=5, usecols='Z:AF', nrows=238, na_values=NA_VALS)
    alm_ppi_raw.columns = list(range(2016, 2023))
    ppi_vals = alm_ppi_raw.values[0]
    rayon_list = list(almaty_inv['rayon'].unique())
    year_list = list(range(2016, 2023))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    almaty_ppi = pd.DataFrame(ppi_records)
    almaty_ppi['ppi'] = clean_col(almaty_ppi['ppi'])

    almaty_merged = outer_merge([almaty_inv, almaty_employment, almaty_controls, almaty_ppi])
    almaty_merged['oblast'] = almaty_merged['oblast'].fillna('Almaty')
    return almaty_merged


def process_karagandy():
    base = Path("../Data/Investments")

    # --- Ulytau investments 2022-2024 ---
    ulytau_rayon_names = ['Zhezkazgan_city', 'Qarazhal_city', 'Satpayev_city', 'Zhanaarka', 'Ulytau']

    uly_all_2022 = pd.read_excel(base / "Ulytau/2022_Investments.xls", sheet_name='13', header=4, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    uly_all_2023 = pd.read_excel(base / "Ulytau/2023_Investments.xlsx", sheet_name='19', header=3, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    uly_all_2024 = pd.read_excel(base / "Ulytau/2024_Investments.xlsx", sheet_name='16', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    uly_non_2022 = pd.read_excel(base / "Ulytau/2022_Investments.xls", sheet_name='14', header=4, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    uly_non_2023 = pd.read_excel(base / "Ulytau/2023_Investments.xlsx", sheet_name='20', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    uly_non_2024 = pd.read_excel(base / "Ulytau/2024_Investments.xlsx", sheet_name='17', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [uly_all_2022, uly_all_2023, uly_all_2024, uly_non_2022, uly_non_2023, uly_non_2024]:
        df.drop(3, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)

    for df in [uly_all_2022, uly_all_2023, uly_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = ulytau_rayon_names
    for df in [uly_non_2022, uly_non_2023, uly_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = ulytau_rayon_names

    uly_2022 = uly_all_2022.merge(uly_non_2022, on='rayon', how='left')
    uly_2023 = uly_all_2023.merge(uly_non_2023, on='rayon', how='left')
    uly_2024 = uly_all_2024.merge(uly_non_2024, on='rayon', how='left')

    # --- Karagandy investments 2022-2024 ---
    kar_rayon_names = ['Qaraghandy_city', 'Balqash_city', 'Priozersk_city', 'Saran_city', 'Temirtau_city', 'Shakhtinsk_city', 'Abay', 'Aqtoghay', 'Buqar_zhyrau', 'Qarqaraly', 'Nura', 'Osakarov', 'Shet']

    kar_all_2022 = pd.read_excel(base / "Karagandy/2022_Investments.xls", sheet_name='13', header=4, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    kar_all_2023 = pd.read_excel(base / "Karagandy/2023_Investments.xls", sheet_name='19', header=3, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    kar_all_2024 = pd.read_excel(base / "Karagandy/2024_Investments.xls", sheet_name='14', header=3, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    kar_non_2022 = pd.read_excel(base / "Karagandy/2022_Investments.xls", sheet_name='14', header=4, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    kar_non_2023 = pd.read_excel(base / "Karagandy/2023_Investments.xls", sheet_name='20', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    kar_non_2024 = pd.read_excel(base / "Karagandy/2024_Investments.xls", sheet_name='15', header=3, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [kar_all_2022, kar_all_2023, kar_all_2024, kar_non_2022, kar_non_2023, kar_non_2024]:
        df.drop(6, inplace=True, errors='ignore')
        df.reset_index(drop=True, inplace=True)

    for df in [kar_all_2022, kar_all_2023, kar_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = kar_rayon_names
    for df in [kar_non_2022, kar_non_2023, kar_non_2024]:
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = kar_rayon_names

    kar_2022 = pd.concat([uly_2022, kar_all_2022.merge(kar_non_2022, on='rayon', how='left')], ignore_index=True)
    kar_2023 = pd.concat([uly_2023, kar_all_2023.merge(kar_non_2023, on='rayon', how='left')], ignore_index=True)
    kar_2024 = pd.concat([uly_2024, kar_all_2024.merge(kar_non_2024, on='rayon', how='left')], ignore_index=True)

    # --- 2014 (camelot) ---
    kar_rayon_names_18 = ['Qaraghandy_city', 'Balqash_city', 'Zhezkazgan_city', 'Qarazhal_city', 'Priozersk_city', 'Saran_city', 'Satpayev_city', 'Temirtau_city', 'Shakhtinsk_city', 'Abay', 'Aqtoghay', 'Buqar_zhyrau', 'Zhanaarka', 'Qarqaraly', 'Nura', 'Osakarov', 'Ulytau', 'Shet']

    t2014 = camelot.read_pdf(Path("../Data/Investments/Karagandy/2014_Investments.pdf"), pages='22', flavor='stream')[0].df
    kar_all_2014 = t2014.iloc[:, [0, 3, 5]].drop([0, 1, 2, 3, 4, 14]).reset_index(drop=True)
    for col in [kar_all_2014.columns[1], kar_all_2014.columns[2]]:
        kar_all_2014[col] = kar_all_2014[col].astype(str).str.replace(' ', '', regex=False).replace('', '0').astype(float)
    kar_all_2014.columns = INV_COLS
    to_numeric_inv(kar_all_2014, ['l_all_inv', 'f_all_inv'])
    kar_all_2014['rayon'] = kar_rayon_names_18
    kar_all_2014['oblast'] = 'Qaraghandy'
    kar_all_2014['year'] = 2014

    # --- 2015-2021 ---
    years_files = {
        2015: (base / "Karagandy/2015_Investments/Internal_vs_external.xls", None, 6, 'A, C, D', base / "Karagandy/2015_Investments/Greenfield_investments.xls", None, 7, 'A, B, C'),
        2016: (base / "Karagandy/2016_Investments/Internal_vs_external.xls", None, 6, 'A, C, E', base / "Karagandy/2016_Investments/Greenfield_investments.xls", None, 7, 'A, B, D'),
        2017: (base / "Karagandy/2017_Investments/Internal_vs_external.xls", None, 6, 'A, C, E', base / "Karagandy/2017_Investments/Greenfield_investments.xls", None, 7, 'A, B, D'),
        2018: (base / "Karagandy/2018_Investments/Internal_vs_external.xls", None, 5, 'A, C, E', base / "Karagandy/2018_Investments/Greenfield_investments.xls", None, 5, 'A, B, D'),
        2019: (base / "Karagandy/2019_Investments.xls", '14', 4, 'A, C, E', base / "Karagandy/2019_Investments.xls", '15', 4, 'A, B, D'),
        2020: (base / "Karagandy/2020_Investments.xls", '13', 4, 'A, C, E', base / "Karagandy/2020_Investments.xls", '14', 4, 'A, B, D'),
        2021: (base / "Karagandy/2021_Investments.xls", '13', 4, 'A, C, E', base / "Karagandy/2021_Investments.xls", '14', 4, 'A, B, D'),
    }

    kar_year_dfs = []
    for yr, (all_path, all_sheet, all_hdr, all_ucols, non_path, non_sheet, non_hdr, non_ucols) in years_files.items():
        kw_all = dict(header=all_hdr, usecols=all_ucols, na_values=NA_VALS)
        if all_sheet:
            kw_all['sheet_name'] = all_sheet
        kw_non = dict(header=non_hdr, usecols=non_ucols, na_values=NA_VALS)
        if non_sheet:
            kw_non['sheet_name'] = non_sheet

        all_df = pd.read_excel(all_path, **kw_all).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(non_path, **kw_non).fillna(0).drop(0).reset_index(drop=True)

        # 2015-2018: greenfield files have extra rows; trim to 18 rayons
        if yr in [2015, 2016, 2017, 2018]:
            non_df = non_df.loc[:17].reset_index(drop=True)
        # For 2019-2021: drop subtotal row 9 (Аудандар subtotal)
        if yr in [2019, 2020, 2021]:
            all_df.drop(9, inplace=True, errors='ignore')
            all_df.reset_index(drop=True, inplace=True)
            non_df.drop(9, inplace=True, errors='ignore')
            non_df.reset_index(drop=True, inplace=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = kar_rayon_names_18
        non_df['rayon'] = kar_rayon_names_18
        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Qaraghandy'
        merged['year'] = yr
        kar_year_dfs.append(merged)

    for df in [kar_2022, kar_2023, kar_2024]:
        df['oblast'] = 'Qaraghandy'
    kar_2022['year'] = 2022
    kar_2023['year'] = 2023
    kar_2024['year'] = 2024

    karagandy_inv = pd.concat([kar_all_2014] + kar_year_dfs + [kar_2022, kar_2023, kar_2024], ignore_index=True)
    karagandy_inv['rayon'] = karagandy_inv['rayon'].str.strip()

    # --- Employment ---
    uly_empl = pd.read_excel(Path("../Data/Employment/Ulytau.xlsx"), skiprows=16, nrows=5, header=None, usecols='A:C', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    uly_empl.columns = ['rayon', 2022, 2023]
    uly_empl['rayon'] = ulytau_rayon_names
    uly_empl_long = uly_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    kar_empl_rayon_map = {
        'Қарағанды қ.': 'Qaraghandy_city', 'Қарағанды қ.ә.': 'Qaraghandy_city',
        'Балқаш қ.': 'Balqash_city', 'Балқаш қ.ә.': 'Balqash_city',
        'Жезқазған қ.': 'Zhezkazgan_city', 'Жезқазған қ.ә.': 'Zhezkazgan_city',
        'Қаражал қ.': 'Qarazhal_city', 'Қаражал қ.ә.': 'Qarazhal_city',
        'Приозер қаласы': 'Priozersk_city', 'Приозерск қ.': 'Priozersk_city',
        'Приозерск қ.ә.': 'Priozersk_city',
        'Саран қ.': 'Saran_city', 'Саран қ.ә.': 'Saran_city',
        'Сәтбаев қ.': 'Satpayev_city', 'Сәтпаев қ.': 'Satpayev_city',
        'Сәтпаев қ.ә.': 'Satpayev_city',
        'Теміртау қ.': 'Temirtau_city', 'Теміртау қ.ә.': 'Temirtau_city',
        'Шахтинск қ.': 'Shakhtinsk_city', 'Шахтинск қ.ә.': 'Shakhtinsk_city',
        'Абай ауданы': 'Abay', 'Абайский': 'Abay',
        'Ақтоғай ауданы': 'Aqtoghay', 'Актогайский': 'Aqtoghay',
        'Бұқар жырау ауданы': 'Buqar_zhyrau', 'Бухар-Жырауский': 'Buqar_zhyrau',
        'Жаңаарқа ауданы': 'Zhanaarka', 'Жанааркинский': 'Zhanaarka',
        'Қарқаралы ауданы': 'Qarqaraly', 'Каркаралинский': 'Qarqaraly',
        'Нұра ауданы': 'Nura', 'Нуринский': 'Nura',
        'Осакаров ауданы': 'Osakarov', 'Осакаровский': 'Osakarov',
        'Шет ауданы': 'Shet', 'Шетский': 'Shet',
        'Ұлытау ауданы': 'Ulytau', 'Улытауский': 'Ulytau',
    }

    kar_empl_22_23 = pd.read_excel(Path("../Data/Employment/Karagandy.xls"), skiprows=30, nrows=18, header=None, usecols='A, J, K', na_values=NA_VALS).dropna().reset_index(drop=True)
    kar_empl_22_23.columns = ['rayon', 2022, 2023]
    kar_empl_22_23['rayon'] = ['Qaraghandy_city', 'Balqash_city', 'Priozersk_city', 'Saran_city', 'Temirtau_city', 'Shakhtinsk_city', 'Abay', 'Aqtoghay', 'Buqar_zhyrau', 'Qarqaraly', 'Nura', 'Osakarov', 'Shet']
    kar_empl_22_23_long = kar_empl_22_23.melt(id_vars='rayon', var_name='year', value_name='employment')

    kar_empl_14_21 = pd.read_excel(Path("../Data/Employment/Karagandy.xls"), skiprows=30, nrows=18, header=None, usecols='A, B, C, D, E, F, G, H, I', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    kar_empl_14_21.columns = ['rayon', 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021]
    kar_empl_14_21['rayon'] = kar_empl_14_21['rayon'].str.strip().replace(kar_empl_rayon_map)
    kar_empl_14_21_long = kar_empl_14_21.melt(id_vars='rayon', var_name='year', value_name='employment')

    karagandy_employment = pd.concat([uly_empl_long, kar_empl_22_23_long, kar_empl_14_21_long], ignore_index=True)

    # --- Controls ---
    ctrl_uly_path = Path("../Data/Controls/Ulytau_All.xls")
    ctrl_kar_path = Path("../Data/Controls/Karagandy_All.xls")

    uly_ctrl_years = list(range(2014, 2024))
    uly_ctrl_sheets = {
        'Zhezkazgan_city': 'г.а.Жезказган',
        'Satpayev_city': 'г.а.Сатпаев',
        'Zhanaarka': 'Жанааркинский район',
        'Ulytau': 'Улытауский район',
    }
    uly_ctrl_dfs = []
    for rayon_name, sheet in uly_ctrl_sheets.items():
        df = ctrl_sheet_to_long(ctrl_uly_path, sheet, 2, 'A, Y:AH', 149, uly_ctrl_years, [71, 87, 140], rayon_name)
        uly_ctrl_dfs.append(df)

    # karazhal (г.а.Каражал) += karazhal_city (г.Каражал)
    karazhal_ctrl = ctrl_sheet_to_long(ctrl_uly_path, 'г.а.Каражал', 2, 'A, Y:AH', 149, uly_ctrl_years, [71, 87, 140], 'Qarazhal_city')
    karazhal_city_ctrl = ctrl_sheet_to_long(ctrl_uly_path, 'г.Каражал', 2, 'A, F:O', 149, uly_ctrl_years, [71, 87, 140], 'Qarazhal_city')
    for yr in uly_ctrl_years:
        karazhal_ctrl.loc[karazhal_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += karazhal_city_ctrl.loc[karazhal_city_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values
    uly_ctrl_dfs.append(karazhal_ctrl)

    # satpayev (г.а.Сатпаев) += satpayev_city (г.Сатпаев)
    satpayev_ctrl = next(d for d in uly_ctrl_dfs if d['rayon'].iloc[0] == 'Satpayev_city')
    satpayev_city_ctrl = ctrl_sheet_to_long(ctrl_uly_path, 'г.Сатпаев', 2, 'A, F:O', 149, uly_ctrl_years, [71, 87, 140], 'Satpayev_city')
    for yr in uly_ctrl_years:
        satpayev_ctrl.loc[satpayev_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += satpayev_city_ctrl.loc[satpayev_city_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values

    kar_ctrl_years = list(range(2014, 2025))
    kar_ctrl_sheets = {
        'Qaraghandy_city': 'Караганда г.а.',
        'Priozersk_city': 'г. Приозерск',
        'Saran_city': 'Сарань г.а.',
        'Abay': 'Абайский',
        'Aqtoghay': 'Актогайский',
        'Buqar_zhyrau': 'Бухар-Жырауский',
        'Qarqaraly': 'Каркаралинский',
        'Nura': 'Нуринский',
        'Osakarov': 'Осакаровский',
        'Shet': 'Шетский',
    }
    kar_ctrl_dfs = []
    for rayon_name, sheet in kar_ctrl_sheets.items():
        df = ctrl_sheet_to_long(ctrl_kar_path, sheet, 2, 'A, Y:AI', 148, kar_ctrl_years, [70, 86, 139], rayon_name)
        kar_ctrl_dfs.append(df)

    # balkhash += balkhash_city; temirtau += temirtau_city; shakhtinsk += shakhtinsk_city
    mono_map = {
        'Balqash_city': ('Балхаш г.а.', 'Балхаш (моногород)', [87, 115, 161]),
        'Temirtau_city': ('Темиртау г.а.', 'Темиртау (моногород)', [87, 115, 161]),
        'Shakhtinsk_city': ('Шахтинск г.а.', 'Шахтинск (моногород)', [88, 116, 162]),
    }
    for rayon_name, (main_sheet, mono_sheet, rows) in mono_map.items():
        main_ctrl = ctrl_sheet_to_long(ctrl_kar_path, main_sheet, 2, 'A, Y:AI', 148, kar_ctrl_years, [70, 86, 139], rayon_name)
        mono_ctrl = ctrl_sheet_to_long(ctrl_kar_path, mono_sheet, 2, 'A, F:P', 164, kar_ctrl_years, rows, rayon_name + '_mono')
        for yr in kar_ctrl_years:
            main_ctrl.loc[main_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += mono_ctrl.loc[mono_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values
        kar_ctrl_dfs.append(main_ctrl)

    # abay += abay_city (monogorod)
    abay_ctrl = next(d for d in kar_ctrl_dfs if d['rayon'].iloc[0] == 'Abay')
    abay_city_ctrl = ctrl_sheet_to_long(ctrl_kar_path, 'Абай (моногород)', 2, 'A, F:P', 164, kar_ctrl_years, [87, 115, 161], 'Abay_mono')
    for yr in kar_ctrl_years:
        abay_ctrl.loc[abay_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += abay_city_ctrl.loc[abay_city_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values

    karagandy_controls = pd.concat(uly_ctrl_dfs + kar_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    kar_ppi_raw = pd.read_excel(Path("../Data/PPI/Karagandy.xls"), header=4, usecols='Y:AI', nrows=216, na_values=NA_VALS)
    kar_ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = kar_ppi_raw.values[75]
    rayon_list = list(karagandy_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    karagandy_ppi = pd.DataFrame(ppi_records)
    karagandy_ppi['ppi'] = clean_col(karagandy_ppi['ppi'])

    karagandy_merged = outer_merge([karagandy_inv, karagandy_employment, karagandy_controls, karagandy_ppi])
    karagandy_merged['oblast'] = karagandy_merged['oblast'].fillna('Qaraghandy')
    return karagandy_merged


def process_aktobe():
    base = Path("../Data/Investments/Aktobe")
    rayon_names = ['Aqtobe_city', 'Alga', 'Aiteke_bi', 'Baighanin', 'Qarghaly', 'Qobda', 'Martok', 'Mughalzhar', 'Oyil', 'Temir', 'Khromtau', 'Shalkar', 'Yrghyz']

    all_files = [
        (base / "2016_Investments/Internal_vs_external.xls", None, 7, 'A, C, E'),
        (base / "2017_Investments/Internal_vs_extrenal.xls", None, 7, 'A, C, E'),
        (base / "2018_Investments/Internal_vs_external.xlsx", None, 7, 'A, C, E'),
        (base / "2019_Investments.xls", '16.', 5, 'A, C, E'),
        (base / "2020_Investments.xls", '15', 5, 'A, C, E'),
        (base / "2021_Investments.xls", '15', 4, 'A, C, E'),
        (base / "2022_Investments.xls", '15', 4, 'A, C, E'),
        (base / "2023_Investments.xlsx", '19', 4, 'A, C, F'),
        (base / "2024_Investments.xlsx", '14', 4, 'A, C, E'),
    ]
    non_files = [
        (base / "2016_Investments/Greenfield_investments.xls", None, 7, 'A, B, D', 14),
        (base / "2017_Investments/Greenfield_investments.xls", None, 7, 'A, B, D', 14),
        (base / "2018_Investments/Greenfield_investments.xlsx", None, 7, 'A, B, D', 14),
        (base / "2019_Investments.xls", '17.', 5, 'A, B, D', None),
        (base / "2020_Investments.xls", '16.', 5, 'A, B, D', None),
        (base / "2021_Investments.xls", '16.', 4, 'A, B, D', None),
        (base / "2022_Investments.xls", '16.', 4, 'A, B, D', None),
        (base / "2023_Investments.xlsx", '20', 4, 'A, B, D', None),
        (base / "2024_Investments.xlsx", '15', 4, 'A, B, D', None),
    ]

    aktobe_year_dfs = []
    for i, ((ap, asht, ahdr, aucols), (np_, nsht, nhdr, nusecols, nnrows)) in enumerate(zip(all_files, non_files)):
        kw_a = dict(header=ahdr, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw_a['sheet_name'] = asht
        kw_n = dict(header=nhdr, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw_n['sheet_name'] = nsht
        if nnrows:
            kw_n['nrows'] = nnrows

        all_df = pd.read_excel(ap, **kw_a).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(np_, **kw_n).fillna(0).drop(0).reset_index(drop=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = rayon_names
        non_df['rayon'] = rayon_names

        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Aqtobe'
        merged['year'] = 2016 + i
        aktobe_year_dfs.append(merged)

    aktobe_inv = pd.concat(aktobe_year_dfs, ignore_index=True)
    aktobe_inv['rayon'] = aktobe_inv['rayon'].str.strip()

    # --- Employment ---
    aktobe_empl = pd.read_excel(Path("../Data/Employment/Aktobe.xlsx"), skiprows=26, nrows=13, header=None, usecols='A, D, E, F, G, H, I, J, K', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    aktobe_empl.columns = ['rayon', 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023]
    aktobe_empl['rayon'] = rayon_names
    aktobe_employment = aktobe_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Aktobe_All.xls")
    ctrl_years = list(range(2016, 2024))
    sheet_rows_map = {
        'Aqtobe_city': ('г. Актобе', [109, 85, 142]),
        'Alga': ('Алгинский', [112, 89, 141]),
        'Aiteke_bi': ('Айтекебийский', [111, 88, 140]),
        'Baighanin': ('Байганинский', [112, 88, 141]),
        'Qarghaly': ('Каргалинский', [111, 88, 140]),
        'Qobda': ('Хобдинский', [113, 90, 142]),
        'Martok': ('Мартукский', [112, 88, 141]),
        'Mughalzhar': ('Мугалжарский', [112, 89, 141]),
        'Oyil': ('Уилский', [112, 89, 141]),
        'Temir': ('Темирский', [111, 88, 141]),
        'Shalkar': ('Шалкарский', [112, 89, 141]),
        'Yrghyz': ('Иргизский', [111, 88, 144]),
    }
    aktobe_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, R:Y', 149, ctrl_years, rows, rayon_name)
        aktobe_ctrl_dfs.append(df)

    # khromtau += khromtau_city
    khromtau_city_ctrl = ctrl_sheet_to_long(ctrl_path, 'г. Хромтау', 2, 'A, AA:AH', 149, ctrl_years, [80, 115, 146], 'Khromtau_city')
    khromtau_ctrl = ctrl_sheet_to_long(ctrl_path, 'Хромтауский', 2, 'A, R:Y', 149, ctrl_years, [112, 89, 141], 'Khromtau')
    for yr in ctrl_years:
        khromtau_ctrl.loc[khromtau_ctrl['year'] == yr, ['production', 'agriculture', 'retail']] += khromtau_city_ctrl.loc[khromtau_city_ctrl['year'] == yr, ['production', 'agriculture', 'retail']].values
    aktobe_ctrl_dfs.append(khromtau_ctrl)

    aktobe_controls = pd.concat(aktobe_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Aktobe.xlsx"), header=2, usecols='AA:AH', nrows=216, na_values=NA_VALS)
    ppi_raw.columns = list(range(2016, 2024))
    ppi_vals = ppi_raw.values[64]
    rayon_list = list(aktobe_inv['rayon'].unique())
    year_list = list(range(2016, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    aktobe_ppi = pd.DataFrame(ppi_records)
    aktobe_ppi['ppi'] = clean_col(aktobe_ppi['ppi'])

    aktobe_merged = outer_merge([aktobe_inv, aktobe_employment, aktobe_controls, aktobe_ppi])
    aktobe_merged['oblast'] = aktobe_merged['oblast'].fillna('Aqtobe')
    return aktobe_merged


def process_almaty_city():
    base = Path("../Data/Investments/Almaty-city")

    rayon_map_alm_city = {
        'Алатау': 'Alatau', 'Алмалы': 'Almaly', 'Әуезов': 'Auezov', 'Бостандық': 'Bostandyq',
        'Жетісу': 'Zhetysu', 'Медеу': 'Medeu', 'Наурызбай': 'Nauryzbay', 'Түрксіб': 'Turksib',
        'Алатауский': 'Alatau', 'Алатауский район': 'Alatau',
        'Алмалинский': 'Almaly', 'Алмалинский район': 'Almaly',
        'Ауэзовский': 'Auezov', 'Ауэзовский район': 'Auezov',
        'Бостандыкский': 'Bostandyq', 'Бостандыкский район': 'Bostandyq',
        'Жетысуский': 'Zhetysu', 'Жетысуский район': 'Zhetysu',
        'Медеуский': 'Medeu', 'Медеуский район': 'Medeu',
        'Наурызбайский': 'Nauryzbay', 'Наурызбайский район': 'Nauryzbay',
        'Турксибский': 'Turksib', 'Турксибский район': 'Turksib', 'Түрксибский': 'Turksib',
        'Алматы қаласы': 'Almaty city', 'г. Алматы': 'Almaty city', 'г.Алматы': 'Almaty city',
    }

    def read_alm_city_all(path, sheet, header, nrows, usecols):
        kw = dict(header=header, nrows=nrows, usecols=usecols, na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_alm_city)
        return df

    alm_city_all_2014 = read_alm_city_all(base / "2014_Investments.xls", None, 7, 9, 'A, C, D')
    alm_city_all_2014 = alm_city_all_2014.replace({'Јуезов': 'Auezov', 'Бостандыќ': 'Bostandyq', 'Тїрксіб': 'Turksib'})
    alm_city_all_2015 = read_alm_city_all(base / "2015_Investments.xls", 'TAB 11', 7, 9, 'A, C, D')
    alm_city_all_2016 = read_alm_city_all(base / "2016_Investments.xls", 'ТАБ 13', 8, 9, 'A, C, E')
    alm_city_all_2017 = read_alm_city_all(base / "2017_Investments.xls", 'ТАВ 14', 7, 9, 'A, C, F')
    alm_city_all_2018 = read_alm_city_all(base / "2018_Investments/Internal_vs_external.xls", None, 7, 9, 'A, C, F')
    alm_city_all_2019 = read_alm_city_all(base / "2019_Investments.xls", '13.', 5, 9, 'A, C, F')
    alm_city_all_2020 = read_alm_city_all(base / "2020_Investments.xls", '12.', 5, 9, 'A, C, F')
    alm_city_all_2021 = read_alm_city_all(base / "2021_Investments.xls", '12.', 5, 9, 'A, C, F')
    alm_city_all_2022 = read_alm_city_all(base / "2022_Investments.xls", '12.', 5, 9, 'A, C, E')
    alm_city_all_2023 = read_alm_city_all(base / "2023_Investments.xlsx", '18', 3, 9, 'A, C, F')
    alm_city_all_2024 = read_alm_city_all(base / "2024_Investments.xlsx", '13', 3, 9, 'A, C, E')

    def read_alm_city_non(path, sheet, header, nrows):
        kw = dict(header=header, nrows=nrows, usecols='A, B, D', na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_alm_city)
        return df

    alm_city_non_2018 = read_alm_city_non(base / "2018_Investments/Greenfield_investments.xls", None, 10, 9)
    alm_city_non_2019 = read_alm_city_non(base / "2019_Investments.xls", '14.', 7, 9)
    alm_city_non_2020 = read_alm_city_non(base / "2020_Investments.xls", '13.', 7, 9)
    alm_city_non_2021 = read_alm_city_non(base / "2021_Investments.xls", '13.', 7, 9)
    alm_city_non_2022 = read_alm_city_non(base / "2022_Investments.xls", '13.', 7, 9)
    alm_city_non_2023 = read_alm_city_non(base / "2023_Investments.xlsx", '19', 3, 9)
    alm_city_non_2024 = read_alm_city_non(base / "2024_Investments.xlsx", '14', 3, 9)

    year_all_dfs = [
        (2014, alm_city_all_2014, None),
        (2015, alm_city_all_2015, None),
        (2016, alm_city_all_2016, None),
        (2017, alm_city_all_2017, None),
        (2018, alm_city_all_2018, alm_city_non_2018),
        (2019, alm_city_all_2019, alm_city_non_2019),
        (2020, alm_city_all_2020, alm_city_non_2020),
        (2021, alm_city_all_2021, alm_city_non_2021),
        (2022, alm_city_all_2022, alm_city_non_2022),
        (2023, alm_city_all_2023, alm_city_non_2023),
        (2024, alm_city_all_2024, alm_city_non_2024),
    ]

    alm_city_year_dfs = []
    for yr, all_df, non_df in year_all_dfs:
        if non_df is not None:
            df = all_df.merge(non_df, on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Almaty-city'
        df['year'] = yr
        alm_city_year_dfs.append(df)

    alm_city_inv = pd.concat(alm_city_year_dfs, ignore_index=True)
    alm_city_inv['rayon'] = alm_city_inv['rayon'].str.strip()

    # --- Employment ---
    alm_city_empl = pd.read_excel(Path("../Data/Employment/Almaty-city.xls"), skiprows=18, nrows=9, header=None, usecols='A, D, E, F, G, H, I, J, K, L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    alm_city_empl.columns = ['rayon', 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024]
    alm_city_empl['rayon'] = alm_city_empl['rayon'].str.strip().replace(rayon_map_alm_city)
    alm_city_empl = alm_city_empl[alm_city_empl['rayon'] != 'Almaty city']
    alm_city_employment = alm_city_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    # Use header=None so 'Алмалинский' is read as data (not consumed as column header)
    agri_df = pd.read_excel(Path("../Data/Controls/Almaty-city_Agriculture.xls"), skiprows=5, nrows=8, usecols='A:J', na_values=NA_VALS, header=None).fillna(0)
    agri_df.columns = ['rayon'] + list(range(2016, 2025))
    agri_df['rayon'] = agri_df['rayon'].astype(str).str.strip().replace(rayon_map_alm_city)
    agri_long = agri_df.melt(id_vars='rayon', var_name='year', value_name='agriculture')

    prod_df = pd.read_excel(Path("../Data/Controls/Almaty-city_Production.xls"), skiprows=5, nrows=8, usecols='A, N:W', na_values=NA_VALS, header=None).fillna(0)
    prod_df.columns = ['rayon'] + list(range(2014, 2024))
    prod_df['rayon'] = prod_df['rayon'].astype(str).str.strip().replace(rayon_map_alm_city)
    prod_long = prod_df.melt(id_vars='rayon', var_name='year', value_name='production')

    retail_df = pd.read_excel(Path("../Data/Controls/Almaty-city_Retail.xlsx"), sheet_name='торговля', skiprows=3, nrows=8, usecols='A, Z:AI', na_values=NA_VALS, header=None).fillna(0)
    retail_df.columns = ['rayon'] + list(range(2014, 2024))
    retail_df['rayon'] = retail_df['rayon'].astype(str).str.strip().replace(rayon_map_alm_city)
    # The source workbook stores the 2017 and 2018 columns as text with comma
    # decimal separators (e.g. '198,7'); clean_col converts them to floats so
    # they don't end up as strings in clean_data.csv.
    for yr in range(2014, 2024):
        retail_df[yr] = clean_col(retail_df[yr])
    retail_long = retail_df.melt(id_vars='rayon', var_name='year', value_name='retail')

    alm_city_controls = outer_merge([agri_long, prod_long, retail_long])

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Almaty-city.xls"), header=3, usecols='X:AH', nrows=8, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[[-4]].values[0]
    rayon_list = list(alm_city_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    alm_city_ppi = pd.DataFrame(ppi_records)
    alm_city_ppi['ppi'] = clean_col(alm_city_ppi['ppi'])

    alm_city_merged = outer_merge([alm_city_inv, alm_city_employment, alm_city_controls, alm_city_ppi])
    alm_city_merged['oblast'] = alm_city_merged['oblast'].fillna('Almaty-city')
    return alm_city_merged


def process_akmola():
    base = Path("../Data/Investments/Akmola")

    rayon_map_akmola = {
        'Көкшетау қ.ә.': 'Kokshetau city', 'Көкшетау қ.': 'Kokshetau city',
        'г. Кокшетау': 'Kokshetau city', 'г.Кокшетау': 'Kokshetau city',
        'г.а. Кокшетау': 'Kokshetau city', 'г.а. Кокшетау ': 'Kokshetau city',
        'Степногор қ.ә.': 'Stepnogorsk city', 'Степногорск қ.ә.': 'Stepnogorsk city',
        'Степногорск қ.': 'Stepnogorsk city',
        'г. Степногорск': 'Stepnogorsk city',
        'г.а. Степногорска': 'Stepnogorsk city',
        'Қосшы қ.ә.': 'Qosshi city', 'Қосшы қ.': 'Qosshi city',
        'г. Қосшы': 'Qosshi city', 'г.а. Косшы': 'Qosshi city',
        'Ақкөл': 'Aqkol', 'Ақкөл ': 'Aqkol', 'Аккольский': 'Aqkol',
        'Аршалы': 'Arshaly', 'Аршалы ': 'Arshaly', 'Аршалынский': 'Arshaly',
        'Астрахан': 'Astrakhan', 'Астрахан ': 'Astrakhan', 'Астраханский': 'Astrakhan',
        'Атбасар': 'Atbasar', 'Атбасар ': 'Atbasar', 'Атбасарский': 'Atbasar',
        'Бұланды': 'Bulandy', 'Бұланды ': 'Bulandy', 'Буландынский': 'Bulandy',
        'Бурабай': 'Burabay', 'Бурабай ': 'Burabay', 'Бурабайский': 'Burabay',
        'Егіндікөл': 'Egindykol', 'Егіндікөл ': 'Egindykol', 'Егиндыкольский': 'Egindykol',
        'Біржан сал': 'Birzhan Sal', 'Биржан сал': 'Birzhan Sal',
        'им. Биржан сала': 'Birzhan Sal',
        'Ереймен': 'Ereymentau', 'Ерейментау ': 'Ereymentau', 'Ерейментауский': 'Ereymentau',
        'Есіл': 'Yesil', 'Есіл ': 'Yesil', 'Есильский': 'Yesil',
        'Жақсы': 'Zhaqsyn', 'Жақсы ': 'Zhaqsyn', 'Жаксынский': 'Zhaqsyn',
        'Жарқайың': 'Zharqain', 'Жарқайың ': 'Zharqain', 'Жаркаинский': 'Zharqain',
        'Зеренді': 'Zerendi', 'Зеренді ': 'Zerendi', 'Зерендинский': 'Zerendi',
        'Қорғалжын': 'Qorghalzhyn', 'Қорғалжын ': 'Qorghalzhyn', 'Коргалжынский': 'Qorghalzhyn',
        'Сандықтау': 'Sandyqtau', 'Сандықтау ': 'Sandyqtau', 'Сандыктауский': 'Sandyqtau',
        'Целиноград': 'Tselinograd', 'Целиноград ': 'Tselinograd', 'Целиноградский': 'Tselinograd',
        'Шортанды': 'Shortandy', 'Шортанды ': 'Shortandy', 'Шортандинский': 'Shortandy',
        # Enbekshilder was renamed to Birzhan Sal in 2017 — map all variants to newer name
        'Еңбекшілдер': 'Birzhan Sal', 'Еңбекшілдер ': 'Birzhan Sal', 'Енбекшильдерский': 'Birzhan Sal',
    }

    rayon_names = ['Kokshetau city', 'Qosshi city', 'Stepnogorsk city', 'Aqkol', 'Arshaly', 'Astrakhan', 'Atbasar', 'Bulandy', 'Egindykol', 'Birzhan Sal', 'Ereymentau', 'Yesil', 'Zhaqsyn', 'Zharqain', 'Zerendi', 'Qorghalzhyn', 'Sandyqtau', 'Tselinograd', 'Shortandy', 'Burabay']

    all_files = [
        (base / "2014_Investments/Internal_vs_external.xls", None, 7, 'A, C, D', 20),
        (base / "2016_Investments/Internal_vs_external.xls", None, 8, 'A, C, E', None),
        (base / "2017_Investments/Internal_vs_external.xls", None, 8, 'A, C, E', None),
        (base / "2018_Investments/Internal_vs_external.xls", None, 6, 'A, C, E', None),
        (base / "2019_Investments.xls", '14', 3, 'A, C, E', None),
        (base / "2020_Investments.xls", '13', 3, 'A, C, E', None),
        (base / "2021_Investments.xls", '13', 3, 'A, C, E', 20),
        (base / "2022_Investments.xlsx", '21', 3, 'A, C, F', 21),
        (base / "2023_Investments.xlsx", '21', 3, 'A, C, F', 21),
        (base / "2024_Investments.xlsx", '14', 3, 'A, C, E', 21),
    ]
    non_files = [
        (base / "2014_Investments/Greenfield_investments.xls", None, 7, 'A, B, C', None),
        (base / "2016_Investments/Greenfield_investments.xls", None, 7, 'A, B, D', None),
        (base / "2017_Investments/Greenfield_investments.xls", None, 7, 'A, B, D', None),
        (base / "2018_Investments/Greenfield_investments.xls", None, 5, 'A, B, D', None),
        (base / "2019_Investments.xls", '15', 3, 'A, B, D', None),
        (base / "2020_Investments.xls", '14', 3, 'A, B, D', None),
        (base / "2021_Investments.xls", '14', 3, 'A, B, D', 21),
        (base / "2022_Investments.xlsx", '22', 3, 'A, B, D', None),
        (base / "2023_Investments.xlsx", '22', 3, 'A, B, D', None),
        (base / "2024_Investments.xlsx", '15', 3, 'A, B, D', None),
    ]
    years = [2014, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024]

    akmola_year_dfs = []
    for yr, (ap, asht, ahdr, aucols, anrows), (np_, nsht, nhdr, nusecols, nnrows) in zip(years, all_files, non_files):
        kw_a = dict(header=ahdr, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw_a['sheet_name'] = asht
        if anrows:
            kw_a['nrows'] = anrows
        kw_n = dict(header=nhdr, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw_n['sheet_name'] = nsht
        if nnrows:
            kw_n['nrows'] = nnrows

        all_df = pd.read_excel(ap, **kw_a).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(np_, **kw_n).fillna(0).drop(0).reset_index(drop=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = all_df['rayon'].replace(rayon_map_akmola)
        non_df['rayon'] = non_df['rayon'].replace(rayon_map_akmola)

        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Aqmola'
        merged['year'] = yr
        akmola_year_dfs.append(merged)

    akmola_inv = pd.concat(akmola_year_dfs, ignore_index=True)
    akmola_inv['rayon'] = akmola_inv['rayon'].str.strip()

    # --- Employment ---
    akmola_empl = pd.read_excel(Path("../Data/Employment/Akmola.xlsx"), skiprows=32, nrows=20, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    akmola_empl.columns = ['rayon'] + list(range(2014, 2025))
    akmola_empl['rayon'] = akmola_empl['rayon'].str.strip().replace(rayon_map_akmola)
    akmola_employment = akmola_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Akmola_All.xlsx")
    ctrl_years = list(range(2014, 2024))
    sheet_map = {
        'Kokshetau city': ('г.а. Кокшетау', [49, 81, 133]),
        'Stepnogorsk city': ('г.а.Степногорск', [49, 81, 133]),
        'Aqkol': ('Аккольский', [49, 81, 133]),
        'Arshaly': ('Аршалынский', [49, 81, 133]),
        'Astrakhan': ('Астраханский', [49, 81, 133]),
        'Atbasar': ('Атбасарский', [49, 81, 133]),
        'Bulandy': ('Буландынский', [49, 81, 133]),
        'Egindykol': ('Егиндыкольский', [49, 81, 133]),
        'Birzhan Sal': ('Биржан сал', [49, 81, 133]),
        'Ereymentau': ('Ерейментауский', [49, 81, 133]),
        'Yesil': ('Есильский', [49, 81, 133]),
        'Zhaqsyn': ('Жаксынский', [49, 81, 133]),
        'Zharqain': ('Жаркаинский', [49, 81, 133]),
        'Zerendi': ('Зерендинский', [49, 81, 133]),
        'Qorghalzhyn': ('Коргалжынский', [49, 81, 133]),
        'Sandyqtau': ('Сандыктауский', [49, 81, 133]),
        'Tselinograd': ('Целиноградский', [49, 81, 133]),
        'Shortandy': ('Шортандинский', [49, 81, 133]),
        'Burabay': ('Бурабайский', [49, 81, 133]),
    }
    akmola_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AH', 149, ctrl_years, rows, rayon_name)
        akmola_ctrl_dfs.append(df)

    # Kosshi city: usecols='A:G', years=[2018..2023]
    kosshi_years = list(range(2018, 2024))
    kosshi_df = ctrl_sheet_to_long(ctrl_path, 'г.а. Косшы', 2, 'A:G', 149, kosshi_years, [49, 81, 133], 'Qosshi city')
    akmola_ctrl_dfs.append(kosshi_df)

    akmola_controls = pd.concat(akmola_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Akmola.xls"), header=2, usecols='Y:AH', nrows=216, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2024))
    ppi_vals = ppi_raw.values[69]
    rayon_list = list(akmola_inv['rayon'].unique())
    year_list = list(range(2014, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    akmola_ppi = pd.DataFrame(ppi_records)
    akmola_ppi['ppi'] = clean_col(akmola_ppi['ppi'])

    akmola_merged = outer_merge([akmola_inv, akmola_employment, akmola_controls, akmola_ppi])
    akmola_merged['oblast'] = akmola_merged['oblast'].fillna('Aqmola')
    return akmola_merged


def process_astana_city():
    base = Path("../Data/Investments/Astana-city")

    rayon_map_astana = {
        'Астана қ.': 'Astana city', 'Астана қаласы': 'Astana city',
        'Нұр-Сұлтан қаласы': 'Astana city',
        'г. Астана': 'Astana city', 'г.Астана': 'Astana city',
        'Нур-Султан қ.': 'Astana city', 'г. Нур-Султан': 'Astana city',
        '"Алматы" ауданы': 'Almaty', 'Алматы ауданы': 'Almaty',
        'Алматинский район': 'Almaty',
        'район Алматы ': 'Almaty', 'район Алматы': 'Almaty',
        '"Есіл" ауданы': 'Yesil', 'Есіл ауданы': 'Yesil',
        'Есильский район': 'Yesil',
        'район Есиль': 'Yesil', 'район Есиль ': 'Yesil',
        '"Сарыарқа" ауданы': 'Saryarqa', 'Сарыарқа ауданы': 'Saryarqa',
        'Сарыаркинский район': 'Saryarqa',
        'район Сарыарка ': 'Saryarqa', 'район Сарыарка': 'Saryarqa',
        '"Байқоңыр" ауданы': 'Baiqonyr', 'Байқоңыр ауданы': 'Baiqonyr',
        'Байконурский район': 'Baiqonyr',
        'район Байконыр ': 'Baiqonyr', 'район Байкоңыр ': 'Baiqonyr',
        'район Байқоңыр': 'Baiqonyr',
        'Нұра ауданы': 'Nura', 'Нуринский район': 'Nura',
        'район Нура ': 'Nura', 'район Нұра ': 'Nura',
    }

    def read_astana_all(path, sheet, header, nrows, usecols):
        kw = dict(header=header, nrows=nrows, usecols=usecols, na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].replace(rayon_map_astana)
        return df

    def read_astana_non(path, sheet, header, nrows):
        kw = dict(header=header, nrows=nrows, usecols='A, B, D', na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].replace(rayon_map_astana)
        return df

    all_specs = [
        (2014, base / "2014_Investments.xls", None, 5, 5, 'A, C:D'),
        (2015, base / "2015_Investments/Internal_vs_external.xls", None, 5, 5, 'A, C:D'),
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 4, 5, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 4, 5, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 4, 6, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '12.', 6, 6, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '11.', 4, 6, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '12.', 4, 6, 'A, C, E'),
        (2022, base / "2022_Investments.xls", '12', 4, 6, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '15', 2, 7, 'A, C, F'),
        (2024, base / "2024_Investments.xls", '13', 2, 7, 'A, C, E'),
    ]
    non_specs = [
        (2015, base / "2015_Investments/Greenfield_investments.xls", None, 5, 5),
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 5, 5),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 5, 5),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 5, 6),
        (2019, base / "2019_Investments.xls", '13.', 6, 6),
        (2020, base / "2020_Investments.xls", '12.', 4, 6),
        (2021, base / "2021_Investments.xls", '13.', 4, 6),
        (2022, base / "2022_Investments.xls", '13', 4, 6),
        (2023, base / "2023_Investments.xlsx", '16', 3, 6),
        (2024, base / "2024_Investments.xls", '14', 2, 7),
    ]

    all_dfs = {}
    for yr, path, sheet, header, nrows, usecols in all_specs:
        all_dfs[yr] = read_astana_all(path, sheet, header, nrows, usecols)

    non_dfs = {}
    for yr, path, sheet, header, nrows in non_specs:
        non_dfs[yr] = read_astana_non(path, sheet, header, nrows)

    astana_year_dfs = []
    for yr, all_df in all_dfs.items():
        if yr in non_dfs:
            df = all_df.merge(non_dfs[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Astana-city'
        df['year'] = yr
        astana_year_dfs.append(df)

    astana_city_inv = pd.concat(astana_year_dfs, ignore_index=True)
    # Nura->Yesil; drop 'Astana city'; groupby sum
    astana_city_inv['rayon'] = astana_city_inv['rayon'].replace({'Nura': 'Yesil'})
    astana_city_inv = astana_city_inv[astana_city_inv['rayon'] != 'Astana city']
    num_cols = [c for c in astana_city_inv.columns if c not in ['rayon', 'oblast', 'year']]
    astana_city_inv = astana_city_inv.groupby(['rayon', 'oblast', 'year'], as_index=False)[num_cols].sum()
    astana_city_inv['rayon'] = astana_city_inv['rayon'].str.strip()

    # --- Employment ---
    # skiprows=17 skips the 'г. Астана' city-total row (consistent with investment drop above)
    astana_empl = pd.read_excel(Path("../Data/Employment/Astana-city.xlsx"), skiprows=17, nrows=4, header=None, usecols='A:K', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    astana_empl.columns = ['rayon'] + list(range(2014, 2024))
    astana_empl['rayon'] = astana_empl['rayon'].str.strip().replace(rayon_map_astana)
    astana_employment = astana_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Astana-city_All.xlsx")
    ctrl_years_full = list(range(2014, 2024))

    astana_ctrl_dfs = []
    ctrl_specs = [
        ('Almaty', 'Алматы', 2, 'A, S:AB', 149, ctrl_years_full, [82, 57, 141]),
        ('Yesil', 'Есиль', 2, 'B, H:R', 149, list(range(2013, 2024)), [82, 57, 141]),
        ('Saryarqa', 'Сарыарка', 2, 'B, S:AB', 149, ctrl_years_full, [82, 57, 141]),
        ('Baiqonyr', 'Байконыр', 2, 'B:H', 149, list(range(2018, 2024)), [82, 57, 141]),
    ]
    for rayon_name, sheet, header, usecols, nrows, years, rows in ctrl_specs:
        df = ctrl_sheet_to_long(ctrl_path, sheet, header, usecols, nrows, years, rows, rayon_name)
        astana_ctrl_dfs.append(df)

    # Nura->Yesil, groupby sum
    nura_ctrl = ctrl_sheet_to_long(ctrl_path, 'Нура', 2, 'B, C', 149, [2023], [82, 57, 141], 'Nura')
    nura_ctrl['rayon'] = 'Yesil'
    astana_ctrl_dfs.append(nura_ctrl)

    astana_controls_raw = pd.concat(astana_ctrl_dfs, ignore_index=True)
    num_ctrl_cols = [c for c in astana_controls_raw.columns if c not in ['rayon', 'year']]
    astana_controls = astana_controls_raw.groupby(['rayon', 'year'], as_index=False)[num_ctrl_cols].sum()
    # drop 2013 rows from Yesil
    astana_controls = astana_controls[astana_controls['year'] != 2013]

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Astana-city.xls"), header=2, usecols='X:AH', nrows=216, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[1].values
    rayon_list = list(astana_city_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    astana_city_ppi = pd.DataFrame(ppi_records)
    astana_city_ppi['ppi'] = clean_col(astana_city_ppi['ppi'])

    astana_city_merged = outer_merge([astana_city_inv, astana_employment, astana_controls, astana_city_ppi])
    astana_city_merged['oblast'] = astana_city_merged['oblast'].fillna('Astana-city')
    return astana_city_merged


def process_atyrau():
    base = Path("../Data/Investments/Atyrau")

    rayon_map_atyrau = {
        'Атырау қ.': 'Atyrau city', 'Атырау қ.ә.': 'Atyrau city',
        'Атырау г.а.': 'Atyrau city', 'Атырауская г.а': 'Atyrau city',
        'Атырау  Г.А.': 'Atyrau city',
        'г. Атырау': 'Atyrau city', 'г.Атырау': 'Atyrau city',
        'Жылыой': 'Zhylyoi', 'Жылыой ': 'Zhylyoi',
        'Жылыойский': 'Zhylyoi', 'Жылыойский район': 'Zhylyoi',
        'Индер': 'Inder', 'Индер ': 'Inder',
        'Индерский': 'Inder', 'Индерский район': 'Inder',
        'Исатай': 'Isatai', 'Исатай ': 'Isatai',
        'Исатайский': 'Isatai', 'Исатайский район': 'Isatai',
        'Қызылқоға': 'Qyzylqoga', 'Қызылқоға ': 'Qyzylqoga',
        'Кызылкогинский': 'Qyzylqoga', 'Кзылкогинский': 'Qyzylqoga',
        'Кзылкогинский район': 'Qyzylqoga', 'Кызылкога': 'Qyzylqoga',
        'Махамбет': 'Makhambet', 'Махамбет ': 'Makhambet',
        'Махамбетский': 'Makhambet', 'Махамбетский район': 'Makhambet',
        'Мақат': 'Maqat', 'Мақат ': 'Maqat',
        'Макатский': 'Maqat', 'Макатский район': 'Maqat', 'Макат': 'Maqat',
        'Құрманғазы': 'Qurmangazy', 'Құрманғазы ': 'Qurmangazy',
        'Курмангазинский': 'Qurmangazy', 'Курмангазинский район': 'Qurmangazy',
        'Курмангазы': 'Qurmangazy',
    }

    all_specs = [
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 6, 9, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 6, 9, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 6, 9, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '14', 4, 9, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13', 4, 9, 'A, C, E'),
        (2021, base / "2021_Investments.xlsx", '13', 3, 9, 'A, C, E'),
        (2022, base / "2022_Investments.xlsx", '14', 3, 9, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '18', 3, 9, 'A, C, F'),
        (2024, base / "2024_Investments.xlsx", '14', 3, 9, 'A, C, E'),
    ]
    non_specs = [
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 7, 9, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 7, 9, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 7, 9, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '15', 4, 9, 'A, B, D'),
        (2020, base / "2020_Investments.xls", '14', 4, 9, 'A, B, D'),
        (2021, base / "2021_Investments.xlsx", '14', 4, 9, 'A, B, D'),
        (2022, base / "2022_Investments.xlsx", '15', 3, 9, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '19', 3, 9, 'A, B, D'),
        (2024, base / "2024_Investments.xlsx", '15', 3, 9, 'A, B, D'),
    ]

    atyrau_year_dfs = []
    for (yr, ap, asht, ahdr, anrows, aucols), (_, np_, nsht, nhdr, nnrows, nusecols) in zip(all_specs, non_specs):
        kw_a = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw_a['sheet_name'] = asht
        kw_n = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw_n['sheet_name'] = nsht

        all_df = pd.read_excel(ap, **kw_a).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(np_, **kw_n).fillna(0).drop(0).reset_index(drop=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = all_df['rayon'].replace(rayon_map_atyrau)
        non_df['rayon'] = non_df['rayon'].replace(rayon_map_atyrau)

        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Atyrau'
        merged['year'] = yr
        atyrau_year_dfs.append(merged)

    atyrau_inv = pd.concat(atyrau_year_dfs, ignore_index=True)
    atyrau_inv['rayon'] = atyrau_inv['rayon'].str.strip()

    # --- Employment ---
    atyrau_empl = pd.read_excel(Path("../Data/Employment/Atyrau.xlsx"), skiprows=21, nrows=8, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    atyrau_empl.columns = ['rayon'] + list(range(2014, 2025))
    atyrau_empl['rayon'] = atyrau_empl['rayon'].str.strip().replace(rayon_map_atyrau)
    atyrau_employment = atyrau_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls (3 separate files) ---
    # skiprows adjusted so the oblast-total row is consumed as pandas header (discarded)
    # and the city-agglomeration row (г.а.) is the first real data row
    prod_df = pd.read_excel(Path("../Data/Controls/Atyrau_Production.xls"), skiprows=5, nrows=8, usecols='A, Y:AI', na_values=NA_VALS).fillna(0)
    prod_df.columns = ['rayon'] + list(range(2013, 2024))
    prod_df['rayon'] = prod_df['rayon'].replace(rayon_map_atyrau)
    prod_long = prod_df.melt(id_vars='rayon', var_name='year', value_name='production')

    agri_df = pd.read_excel(Path("../Data/Controls/Atyrau_Agriculture.xls"), skiprows=4, nrows=8, usecols='A, U:AE', na_values=NA_VALS).fillna(0)
    agri_df.columns = ['rayon'] + list(range(2013, 2024))
    agri_df['rayon'] = agri_df['rayon'].replace(rayon_map_atyrau)
    agri_long = agri_df.melt(id_vars='rayon', var_name='year', value_name='agriculture')

    retail_df = pd.read_excel(Path("../Data/Controls/Atyrau_Retail.xls"), skiprows=4, nrows=8, usecols='A, O:Y', na_values=NA_VALS).fillna(0)
    retail_df.columns = ['rayon'] + list(range(2013, 2024))
    retail_df['rayon'] = retail_df['rayon'].replace(rayon_map_atyrau)
    for col in [c for c in retail_df.columns if c != 'rayon']:
        retail_df[col] = clean_col(retail_df[col])
    retail_long = retail_df.melt(id_vars='rayon', var_name='year', value_name='retail')

    atyrau_controls = outer_merge([prod_long, agri_long, retail_long])

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Atyrau.xls"), header=2, usecols='X:AH', nrows=65, na_values=NA_VALS)
    ppi_raw.columns = list(range(2013, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(atyrau_inv['rayon'].unique())
    year_list = list(range(2013, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    atyrau_ppi = pd.DataFrame(ppi_records)
    atyrau_ppi['ppi'] = clean_col(atyrau_ppi['ppi'])

    atyrau_merged = outer_merge([atyrau_inv, atyrau_employment, atyrau_controls, atyrau_ppi])
    atyrau_merged['oblast'] = atyrau_merged['oblast'].fillna('Atyrau')
    return atyrau_merged


def process_kostanay():
    base = Path("../Data/Investments/Kostanay")

    rayon_map_kostanay = {
        'Қостанай қ.': 'Qostanay city', 'Қостанай қаласы': 'Qostanay city',
        'г. Костанай': 'Qostanay city', 'г. Костанай ': 'Qostanay city', 'г.Костанай': 'Qostanay city',
        'Костанайский ': 'Qostanay city',
        'Арқалық қ.': 'Arqalyq city', 'Арқалық қ.ә.': 'Arqalyq city',
        'Аркалык г.а.': 'Arqalyq city', 'Аркалык г.а. ': 'Arqalyq city',
        'г. Аркалык': 'Arqalyq city',
        'Лисаковск қ.': 'Lisakovsk city', 'Лисаков қ.ә.': 'Lisakovsk city',
        'Лисаков г.а.': 'Lisakovsk city', 'Лисаковск г.а.': 'Lisakovsk city',
        'г. Лисаковск': 'Lisakovsk city',
        'Рудный қ.': 'Rudny city', 'Рудный қ. ә.': 'Rudny city',
        'Рудный қ.ә.': 'Rudny city', 'Рудный г.а.': 'Rudny city',
        'г. Рудный': 'Rudny city',
        'Алтынсарин': 'Altynsarin', 'Алтынсаринский': 'Altynsarin',
        'Алтынсаринский ': 'Altynsarin',
        'Амангелді': 'Amangeldi', 'Амангельдинский': 'Amangeldi',
        'Амангельдинский ': 'Amangeldi',
        'Әулиекөл': 'Auliekol', 'Ауликольский': 'Auliekol',
        'Аулиекольский': 'Auliekol', 'Аулиекольский ': 'Auliekol',
        'Денисов': 'Denisov', 'Денисовский': 'Denisov', 'Денисовский ': 'Denisov',
        'Жангелді': 'Zhangeldi', 'Жангелдин': 'Zhangeldi',
        'Джангельдинский': 'Zhangeldi', 'Джангельдинский ': 'Zhangeldi',
        'Жітіқара': 'Zhitiqara', 'Житикаринский': 'Zhitiqara', 'Житикаринский ': 'Zhitiqara',
        'Қамысты': 'Qamysty', 'Камыстинский': 'Qamysty', 'Камыстинский ': 'Qamysty',
        'Қарабалық': 'Qarabalyq', 'Карабалыкский': 'Qarabalyq', 'Карабалыкский ': 'Qarabalyq',
        'Қарасу': 'Qarasu', 'Карасуский': 'Qarasu', 'Карасуский ': 'Qarasu',
        'Қостанай ауданы': 'Qostanay', 'Қостанай': 'Qostanay',
        'Костанайский': 'Qostanay', 'Костанайский ': 'Qostanay',
        'Меңдіқара': 'Mendiqara', 'Мендыкаринский': 'Mendiqara', 'Мендыкаринский ': 'Mendiqara',
        'Науырзым': 'Naurzum', 'Наурзумский': 'Naurzum', 'Наурзумский ': 'Naurzum',
        'Сарыкөл': 'Saryqol', 'Сарыкольский': 'Saryqol', 'Сарыкольский ': 'Saryqol',
        'Сұлтанбек Қожанов': 'Sultanbek Qozhanov', 'Ұзынкөл': 'Sultanbek Qozhanov',
        'Узункольский': 'Sultanbek Qozhanov', 'Узункольский ': 'Sultanbek Qozhanov',
        # Taran was renamed to Beimbet Mailin in 2019 — map all variants to newer name
        'Таран': 'Beimbet Mailin', 'Тарановский': 'Beimbet Mailin',
        'Федоров': 'Fyodorov', 'Федоровский': 'Fyodorov', 'Федоровский ': 'Fyodorov',
        'Бейiмбет Майлин': 'Beimbet Mailin', 'Беимбета Майлина': 'Beimbet Mailin',
    }

    all_specs = [
        (2014, base / "2014_Investments/Internal_vs_external.xls", None, 6, 21, 'A, C, D'),
        (2015, base / "2015_Investments/Internal_vs_external.xls", None, 6, 21, 'A, C, D'),
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '13.', 4, 21, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13.', 4, 21, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13.', 4, 21, 'A, C, E'),
        (2022, base / "2022_Investments.xls", '13.', 4, 21, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '19', 3, 21, 'A, C, F'),
        (2024, base / "2024_Investments.xlsx", '14', 3, 21, 'A, C, E'),
    ]
    non_specs = [
        (2014, base / "2014_Investments/Greenfield_investments.xls", None, 6, 21, 'A, B, C'),
        (2015, base / "2015_Investments/Greenfield_investments.xls", None, 6, 21, 'A, B, C'),
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 5, 21, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 5, 21, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 5, 21, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '14.', 5, 21, 'A, B, D'),
        (2020, base / "2020_Investments.xls", '14.', 5, 21, 'A, B, D'),
        (2021, base / "2021_Investments.xls", '14.', 5, 21, 'A, B, D'),
        (2022, base / "2022_Investments.xls", '14.', 5, 21, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '20', 3, 21, 'A, B, D'),
        (2024, base / "2024_Investments.xlsx", '15', 3, 21, 'A, B, D'),
    ]

    kostanay_year_dfs = []
    for (yr, ap, asht, ahdr, anrows, aucols), (_, np_, nsht, nhdr, nnrows, nusecols) in zip(all_specs, non_specs):
        kw_a = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw_a['sheet_name'] = asht
        kw_n = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw_n['sheet_name'] = nsht

        all_df = pd.read_excel(ap, **kw_a).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(np_, **kw_n).fillna(0).drop(0).reset_index(drop=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = all_df['rayon'].replace(rayon_map_kostanay)
        non_df['rayon'] = non_df['rayon'].replace(rayon_map_kostanay)

        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Qostanay'
        merged['year'] = yr
        kostanay_year_dfs.append(merged)

    kostanay_inv = pd.concat(kostanay_year_dfs, ignore_index=True)
    kostanay_inv['rayon'] = kostanay_inv['rayon'].str.strip()

    # --- Employment ---
    kostanay_empl = pd.read_excel(Path("../Data/Employment/Kostanay.xlsx"), skiprows=31, nrows=20, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    kostanay_empl.columns = ['rayon'] + list(range(2014, 2025))
    kostanay_empl['rayon'] = kostanay_empl['rayon'].str.strip().replace(rayon_map_kostanay)
    kostanay_employment = kostanay_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Kostanay_All.xlsx")
    ctrl_years = list(range(2014, 2024))
    sheet_rows_map = {
        'Qostanay city': ('г. Костанай', [79, 108, 151]),
        'Arqalyq city': ('Аркалык г.а.', [79, 107, 150]),
        'Lisakovsk city': ('Лисаковск г.а.', [79, 107, 150]),
        'Rudny city': ('Рудный  г.а.', [79, 107, 150]),
        'Altynsarin': ('Алтынсаринский', [65, 91, 134]),
        'Amangeldi': ('Амангельдинский', [65, 91, 134]),
        'Auliekol': ('Аулиекольский', [65, 91, 134]),
        'Denisov': ('Денисовский', [65, 91, 134]),
        'Zhangeldi': ('Джангельдинский', [65, 91, 134]),
        'Zhitiqara': ('Житикаринский', [65, 91, 134]),
        'Qamysty': ('Камыстинский', [65, 91, 134]),
        'Qarabalyq': ('Карабалыкский', [65, 91, 134]),
        'Qarasu': ('Карасуский', [65, 91, 134]),
        'Qostanay': ('Костанайский', [65, 93, 136]),
        'Mendiqara': ('Мендыкаринский', [65, 91, 134]),
        'Naurzum': ('Наурзумский', [65, 91, 134]),
        'Saryqol': ('Сарыкольский', [65, 91, 134]),
        'Sultanbek Qozhanov': ('Узункольский', [65, 91, 134]),
        'Beimbet Mailin': ('Беимбета Майлина', [65, 91, 134]),
        'Fyodorov': ('Федоровский', [64, 90, 133]),
    }
    kostanay_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AH', 160, ctrl_years, rows, rayon_name)
        kostanay_ctrl_dfs.append(df)

    kostanay_controls = pd.concat(kostanay_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Kostanay.xlsx"), header=2, usecols='X:AH', nrows=76, na_values=NA_VALS)
    ppi_raw.columns = list(range(2013, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(kostanay_inv['rayon'].unique())
    year_list = list(range(2013, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    kostanay_ppi = pd.DataFrame(ppi_records)
    kostanay_ppi['ppi'] = clean_col(kostanay_ppi['ppi'])
    # drop year==2013
    kostanay_ppi = kostanay_ppi[kostanay_ppi['year'] != 2013]

    kostanay_merged = outer_merge([kostanay_inv, kostanay_employment, kostanay_controls, kostanay_ppi])
    kostanay_merged['oblast'] = kostanay_merged['oblast'].fillna('Qostanay')
    return kostanay_merged


def process_kyzylorda():
    base = Path("../Data/Investments/Kyzylorda")

    rayon_map_kyzylorda = {
        'Қызылорда қ.ә.': 'Qyzylorda city', 'Қызылорда қ.': 'Qyzylorda city',
        'Қызылорда г.а.': 'Qyzylorda city', 'Кызылорда г.а.': 'Qyzylorda city',
        'г. Кызылорда': 'Qyzylorda city', 'г.Кызылорда': 'Qyzylorda city',
        'Байқоңыр қ.': 'Baiqonyr city', 'Байқоңыр қ.ә.': 'Baiqonyr city',
        'Байконыр г.а.': 'Baiqonyr city', 'Байконыр г.а': 'Baiqonyr city',
        'г. Байконур': 'Baiqonyr city',
        'Арал ауданы': 'Aral', 'Арал': 'Aral', 'Арал ': 'Aral', 'Аральский': 'Aral',
        'Жалағаш ауданы': 'Zhalagash', 'Жалағаш': 'Zhalagash', 'Жалағаш ': 'Zhalagash',
        'Жалагашский': 'Zhalagash', 'Жалагашский ': 'Zhalagash',
        'Жаңақорған ауданы': 'Zhanaqorgan', 'Жаңақорған': 'Zhanaqorgan',
        'Жаңақорған ': 'Zhanaqorgan', 'Жанакорганский': 'Zhanaqorgan',
        'Жанакорганский ': 'Zhanaqorgan',
        'Қазалы ауданы': 'Qazaly', 'Қазалы': 'Qazaly', 'Қазалы ': 'Qazaly',
        'Казалинский': 'Qazaly', 'Казалинский ': 'Qazaly',
        'Қармақшы ауданы': 'Qarmaqshy', 'Қармақшы': 'Qarmaqshy', 'Қармақшы ': 'Qarmaqshy',
        'Кармакшинский': 'Qarmaqshy', 'Кармакшинский ': 'Qarmaqshy',
        'Сырдария ауданы': 'Syrdarya', 'Сырдария': 'Syrdarya', 'Сырдария ': 'Syrdarya',
        'Сырдарьинский': 'Syrdarya', 'Сырдарийский': 'Syrdarya', 'Сырдарьинский ': 'Syrdarya',
        'Шиелі ауданы': 'Shieli', 'Шиелі': 'Shieli', 'Шиелі ': 'Shieli',
        'Шиелийский': 'Shieli',
    }

    all_specs = [
        (2014, base / "2014_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2015, base / "2015_Investments/Internal_vs_external.xls", None, 6, 21, 'A, C, D'),
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 5, 21, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '14.', 5, 21, 'A, C, E'),
        (2020, base / "2020_Investments.xlsx", '13.', 5, 21, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13.', 3, 21, 'A, C, E'),
        (2022, base / "2022_Investments.xlsx", '13', 3, 21, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '19', 3, 21, 'A, C, F'),
        (2024, base / "2024_Investments.xlsx", '14', 4, 21, 'A, C, E'),
    ]
    non_specs = [
        (2014, base / "2014_Investments/Greenfield_investments.xls", None, 5, None, 'A, B, F'),
        (2015, base / "2015_Investments/Greenfield_investments.xls", None, 5, 10, 'A, B, C'),
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 5, 10, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 5, 10, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 5, 10, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '15.', 5, 10, 'A, B, D'),
        (2020, base / "2020_Investments.xlsx", '14.', 5, 10, 'A, B, D'),
        (2021, base / "2021_Investments.xls", '14.', 4, 10, 'A, B, D'),
        (2022, base / "2022_Investments.xlsx", '14', 3, 10, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '20', 3, 10, 'A, B, D'),
        (2024, base / "2024_Investments.xlsx", '15', 4, 10, 'A, B, D'),
    ]

    kyzylorda_year_dfs = []
    for (yr, ap, asht, ahdr, anrows, aucols), (_, np_, nsht, nhdr, nnrows, nusecols) in zip(all_specs, non_specs):
        kw_a = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw_a['sheet_name'] = asht
        kw_n = dict(header=nhdr, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw_n['sheet_name'] = nsht
        if nnrows:
            kw_n['nrows'] = nnrows

        all_df = pd.read_excel(ap, **kw_a).fillna(0).drop(0).reset_index(drop=True)
        non_df = pd.read_excel(np_, **kw_n).fillna(0).drop(0).reset_index(drop=True)

        all_df.columns = INV_COLS
        to_numeric_inv(all_df, ['l_all_inv', 'f_all_inv'])
        non_df.columns = NON_COLS
        to_numeric_inv(non_df, ['all_nonextr_inv', 'f_nonextr_inv'])
        non_df['l_nonextr_inv'] = non_df['all_nonextr_inv'] - non_df['f_nonextr_inv']
        non_df.drop('all_nonextr_inv', axis=1, inplace=True)
        all_df['rayon'] = all_df['rayon'].replace(rayon_map_kyzylorda)
        non_df['rayon'] = non_df['rayon'].replace(rayon_map_kyzylorda)

        merged = all_df.merge(non_df, on='rayon', how='left')
        merged['oblast'] = 'Qyzylorda'
        merged['year'] = yr
        kyzylorda_year_dfs.append(merged)

    kyzylorda_inv = pd.concat(kyzylorda_year_dfs, ignore_index=True)
    kyzylorda_inv['rayon'] = kyzylorda_inv['rayon'].str.strip()

    # --- Employment ---
    kyzylorda_empl = pd.read_excel(Path("../Data/Employment/Kyzylorda.xls"), skiprows=4, nrows=9, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    kyzylorda_empl.columns = ['rayon'] + list(range(2014, 2025))
    kyzylorda_empl['rayon'] = kyzylorda_empl['rayon'].str.strip().replace(rayon_map_kyzylorda)
    kyzylorda_employment = kyzylorda_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Kyzylorda_All.xls")
    ctrl_years = list(range(2014, 2024))
    sheet_rows_map = {
        'Qyzylorda city': ('г.Кызылорда', [82, 109, 160]),
        'Aral': ('Арал', [83, 108, 159]),
        'Zhalagash': ('Жалагаш', [83, 108, 159]),
        'Zhanaqorgan': ('Жанакорган', [83, 108, 159]),
        'Qazaly': ('Казалы', [83, 108, 159]),
        'Qarmaqshy': ('Кармакшы', [82, 107, 158]),
        'Syrdarya': ('Сырдария', [82, 107, 158]),
        'Shieli': ('Шиели', [82, 108, 159]),
        'Baiqonyr city': ('г.Байконыр', [79, 131, 129]),
    }
    kyzylorda_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AH', 180, ctrl_years, rows, rayon_name)
        kyzylorda_ctrl_dfs.append(df)

    kyzylorda_controls = pd.concat(kyzylorda_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Kyzylorda.xls"), header=2, usecols='Y:AH', nrows=79, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(kyzylorda_inv['rayon'].unique())
    year_list = list(range(2014, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    kyzylorda_ppi = pd.DataFrame(ppi_records)
    kyzylorda_ppi['ppi'] = clean_col(kyzylorda_ppi['ppi'])

    kyzylorda_merged = outer_merge([kyzylorda_inv, kyzylorda_employment, kyzylorda_controls, kyzylorda_ppi])
    kyzylorda_merged['oblast'] = kyzylorda_merged['oblast'].fillna('Qyzylorda')
    return kyzylorda_merged


def process_mangystau():
    base = Path("../Data/Investments/Mangystau")

    rayon_map_mangystau = {
        'Ақтау қ.ә.': 'Aqtau city', 'Ақтау қалалық әкімшілігі': 'Aqtau city',
        'г.Актау': 'Aqtau city', 'Актау городская администрация': 'Aqtau city', 'Актау г. а.': 'Aqtau city',
        'Жаңаөзен қ.ә.': 'Zhanaozen city', 'Жаңаөзен қалалық әкімшілігі': 'Zhanaozen city',
        'г.Жанаозен': 'Zhanaozen city', 'Жанаозен городская администрация': 'Zhanaozen city', 'Жанаозен г. а.': 'Zhanaozen city',
        'Бейнеу ауданы': 'Beineu', 'Бейнеуский район': 'Beineu',
        'Қарақия ауданы': 'Qaraqiyan', 'Каракиянский район': 'Qaraqiyan',
        'Маңғыстау ауданы': 'Mangystau', 'Мангистауский район': 'Mangystau', 'Мангистауская область': 'Mangystau',
        'Мұнайлы ауданы': 'Munaily', 'Мунайлинский район': 'Munaily',
        'Түпқараған ауданы': 'Tupqaragan', 'Тупкараганский район': 'Tupqaragan',
    }

    all_specs = [
        (2014, base / "2014_Investments.xls", None, 5, 8, 'A, C, E'),
        (2015, base / "2015_Investments.xls", None, 5, 8, 'A, C, E'),
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 5, 8, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 3, 8, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 4, 8, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '14', 4, 8, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13', 4, 8, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13', 4, 8, 'A, C, E'),
        (2022, base / "2022_Investments.xlsx", '12', 3, 8, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '19', 3, 8, 'A, C, F'),
        (2024, base / "2024_Investments.xlsx", '14', 3, 8, 'A, C, E'),
    ]
    non_specs = [
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 6, 8, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 4, 8, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 4, 8, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '15', 4, 8, 'A, B, D'),
        (2020, base / "2020_Investments.xls", '14', 4, 8, 'A, B, D'),
        (2021, base / "2021_Investments.xls", '14', 4, 8, 'A, B, D'),
        (2022, base / "2022_Investments.xlsx", '13', 3, 10, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '20', 3, 10, 'A, B, D'),
        (2024, base / "2024_Investments.xlsx", '15', 3, 8, 'A, B, D'),
    ]

    all_dfs = {}
    for yr, ap, asht, ahdr, anrows, aucols in all_specs:
        kw = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw['sheet_name'] = asht
        df = pd.read_excel(ap, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].replace(rayon_map_mangystau)
        all_dfs[yr] = df

    non_dfs = {}
    for yr, np_, nsht, nhdr, nnrows, nusecols in non_specs:
        kw = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw['sheet_name'] = nsht
        df = pd.read_excel(np_, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].replace(rayon_map_mangystau)
        non_dfs[yr] = df

    mangystau_year_dfs = []
    for yr in sorted(all_dfs.keys()):
        all_df = all_dfs[yr]
        if yr in non_dfs:
            df = all_df.merge(non_dfs[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Mangystau'
        df['year'] = yr
        mangystau_year_dfs.append(df)

    mangystau_inv = pd.concat(mangystau_year_dfs, ignore_index=True)
    mangystau_inv['rayon'] = mangystau_inv['rayon'].str.strip()

    # --- Employment ---
    mangystau_empl = pd.read_excel(Path("../Data/Employment/Mangystau.xls"), skiprows=18, nrows=7, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    mangystau_empl.columns = ['rayon'] + list(range(2014, 2025))
    mangystau_empl['rayon'] = mangystau_empl['rayon'].str.strip().replace(rayon_map_mangystau)
    mangystau_employment = mangystau_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Mangystau_All.xlsx")
    ctrl_years = list(range(2014, 2024))
    sheet_rows_map = {
        'Aqtau city': ('г.а.Актау', 'B, Z:AI', [74, 53, 109]),
        'Zhanaozen city': ('г.а. Жанаозен', 'B, Z:AI', [74, 53, 109]),
        'Beineu': ('Бейнеуский район', 'B, Z:AI', [74, 53, 109]),
        'Qaraqiyan': ('Каракиянский район', 'B, Z:AI', [74, 53, 109]),
        'Mangystau': ('Мангистауский район', 'B, Z:AI', [74, 53, 109]),
        'Munaily': ('Мунайлинский район', 'B, Q:Z', [74, 53, 109]),
        'Tupqaragan': ('Тупкараганский район', 'B, Z:AI', [74, 53, 109]),
    }
    mangystau_ctrl_dfs = []
    for rayon_name, (sheet, ucols, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, ucols, 180, ctrl_years, rows, rayon_name)
        mangystau_ctrl_dfs.append(df)

    mangystau_controls = pd.concat(mangystau_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Mangystau.xls"), header=2, usecols='Y:AH', nrows=65, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(mangystau_inv['rayon'].unique())
    year_list = list(range(2014, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    mangystau_ppi = pd.DataFrame(ppi_records)
    mangystau_ppi['ppi'] = clean_col(mangystau_ppi['ppi'])

    mangystau_merged = outer_merge([mangystau_inv, mangystau_employment, mangystau_controls, mangystau_ppi])
    mangystau_merged['oblast'] = mangystau_merged['oblast'].fillna('Mangystau')
    return mangystau_merged


def process_north_kazakhstan():
    base = Path("../Data/Investments/North-Kazakhstan")

    rayon_map_north = {
        'Петропавл қ.': 'Petropavl city', 'Петропавл қаласы': 'Petropavl city',
        'г.Петропавловск': 'Petropavl city', 'Петропавловск': 'Petropavl city',
        'Айыртау ауданы': 'Aiyrtau', 'Айыртауский район': 'Aiyrtau', 'Айыртауский': 'Aiyrtau',
        'Ақжар ауданы': 'Aqzhar', 'Акжарский район': 'Aqzhar', 'Акжарский': 'Aqzhar',
        'Мағжан Жұмабаев ауданы': 'Magzhan Zhumabayev', 'район Магжана Жумабаева': 'Magzhan Zhumabayev', 'М.Жумабаева': 'Magzhan Zhumabayev',
        'Есіл ауданы': 'Esil', 'Есильский район': 'Esil', 'Есильский': 'Esil',
        'Жамбыл ауданы': 'Zhambyl', 'Жамбылский район': 'Zhambyl', 'Жамбылский': 'Zhambyl',
        'Қызылжар ауданы': 'Qyzylzhar', 'Кызылжарский район': 'Qyzylzhar', 'Кызылжарский': 'Qyzylzhar',
        'Мамлют ауданы': 'Mamlyut', 'Мамлютский район': 'Mamlyut', 'Мамлютский': 'Mamlyut',
        'Шал ақын ауданы': 'Shal Aqyn', 'Шал Ақын ауданы': 'Shal Aqyn', 'район Шал акына': 'Shal Aqyn', 'Шал акына': 'Shal Aqyn',
        'Аққайың ауданы': 'Aqqayin', 'Аккайынский район': 'Aqqayin', 'Аккайынский': 'Aqqayin',
        'Тайынша ауданы': 'Taiynsha', 'Тайыншинский район': 'Taiynsha', 'Тайыншинский': 'Taiynsha',
        'Тимирязев ауданы': 'Timiryazev', 'Тимирязевский район': 'Timiryazev', 'Тимирязевский': 'Timiryazev',
        'Уәлиханов ауданы': 'Ualikhanov', 'Уалихановский район': 'Ualikhanov', 'Уалихановкий': 'Ualikhanov',
        'Ғабит Мүсірепов ауданы': 'Gabit Musrepov', 'район им.Габита Мусрепова': 'Gabit Musrepov', 'район им. Габита Мусрепова': 'Gabit Musrepov', 'Г.Мусрепова': 'Gabit Musrepov',
    }

    # 2015-2018 PDF via camelot
    def read_north_pdf_all(pages, cols, drop_rows, col_indices):
        t = camelot.read_pdf(base / f"20{pages[-2:] if len(pages)==2 else pages[:2]}_Investments.pdf", pages=pages, flavor='stream')[0].df
        df = t.iloc[:, col_indices].drop(drop_rows).reset_index(drop=True)
        for c in df.columns[1:]:
            df[c] = df[c].astype(str).str.replace(' ', '', regex=False).astype(int)
        return df

    def read_north_pdf_generic(path, pages, row_start, col_indices):
        t = camelot.read_pdf(path, pages=pages, flavor='stream')[0].df
        df = t.iloc[row_start:, col_indices].reset_index(drop=True)
        return df

    north_all_2015_raw = camelot.read_pdf(base / "2015_Investments.pdf", pages='19', flavor='stream')[0].df
    north_all_2015 = north_all_2015_raw.iloc[6:, [0, 3, 5]].reset_index(drop=True)
    for col in [north_all_2015.columns[1], north_all_2015.columns[2]]:
        north_all_2015[col] = north_all_2015[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    north_all_2016_raw = camelot.read_pdf(base / "2016_Investments.pdf", pages='20', flavor='stream')[0].df
    north_all_2016 = north_all_2016_raw.iloc[6:, [0, 3, 5]].reset_index(drop=True)
    for col in [north_all_2016.columns[1], north_all_2016.columns[2]]:
        north_all_2016[col] = north_all_2016[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    north_all_2017_raw = camelot.read_pdf(base / "2017_Investments.pdf", pages='23', flavor='stream')[0].df
    north_all_2017 = north_all_2017_raw.iloc[6:, [0, 3, 5]].reset_index(drop=True)
    for col in [north_all_2017.columns[1], north_all_2017.columns[2]]:
        north_all_2017[col] = north_all_2017[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    # 2018 special handling
    north_all_2018_raw = camelot.read_pdf(base / "2018_Investments.pdf", pages='29', flavor='stream', table_areas=['0,800,800,0'])[0].df
    df2018 = north_all_2018_raw.copy()
    df2018['col_new_even'] = pd.NA
    even_rows = df2018.index[df2018.index % 2 == 0]
    shifted_indices = even_rows + 1
    shifted_indices = shifted_indices[shifted_indices < len(df2018)]
    df2018.loc[shifted_indices, 'col_new_even'] = df2018.loc[even_rows[:len(shifted_indices)], 0].values
    df2018 = df2018[df2018.index % 2 == 1].reset_index(drop=True)
    df2018 = df2018[['col_new_even', df2018.columns[1], df2018.columns[4]]]
    df2018 = df2018.iloc[10:].reset_index(drop=True)
    for col in [df2018.columns[1], df2018.columns[2]]:
        df2018[col] = df2018[col].astype(str).str.replace(' ', '', regex=False).astype(int)
    north_all_2018 = df2018

    north_all_2019 = pd.read_excel(base / "2019_Investments.xls", sheet_name='14', header=3, nrows=15, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_all_2020 = pd.read_excel(base / "2020_Investments.xls", sheet_name='13', header=3, nrows=15, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_all_2021 = pd.read_excel(base / "2021_Investments.xls", sheet_name='13', header=3, nrows=15, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_all_2022 = pd.read_excel(base / "2022_Investments.xls", sheet_name='13', header=3, nrows=15, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_all_2023 = pd.read_excel(base / "2023_Investments.xlsx", sheet_name='20', header=3, nrows=15, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_all_2024 = pd.read_excel(base / "2024_Investments.xls", sheet_name='14', header=3, nrows=15, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    # non
    north_non_2016_raw = camelot.read_pdf(base / "2016_Investments.pdf", pages='21', flavor='stream')[0].df
    north_non_2016 = north_non_2016_raw.iloc[4:, [0, 2, 6]].reset_index(drop=True)
    for col in [north_non_2016.columns[1], north_non_2016.columns[2]]:
        north_non_2016[col] = north_non_2016[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    north_non_2017_raw = camelot.read_pdf(base / "2017_Investments.pdf", pages='24', flavor='stream')[0].df
    north_non_2017 = north_non_2017_raw.iloc[4:, [0, 2, 6]].reset_index(drop=True)
    for col in [north_non_2017.columns[1], north_non_2017.columns[2]]:
        north_non_2017[col] = north_non_2017[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    north_non_2018_raw = camelot.read_pdf(base / "2018_Investments.pdf", pages='30', flavor='stream')[0].df
    dfs_non = north_non_2018_raw.copy()
    dfs_non[0] = dfs_non[0].shift(1)
    dfs_non = dfs_non.dropna(subset=[0]).reset_index(drop=True)
    dfs_non = dfs_non[dfs_non.index % 2 == 1]
    dfs_non = dfs_non.iloc[2:].reset_index(drop=True)
    dfs_non = dfs_non[[dfs_non.columns[0], dfs_non.columns[1], dfs_non.columns[5]]]
    north_non_2018 = dfs_non

    north_non_2019 = pd.read_excel(base / "2019_Investments.xls", sheet_name='15', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_non_2020 = pd.read_excel(base / "2020_Investments.xls", sheet_name='14', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_non_2021 = pd.read_excel(base / "2021_Investments.xls", sheet_name='14', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_non_2022 = pd.read_excel(base / "2022_Investments.xls", sheet_name='14', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_non_2023 = pd.read_excel(base / "2023_Investments.xlsx", sheet_name='21', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)
    north_non_2024 = pd.read_excel(base / "2024_Investments.xls", sheet_name='15', header=3, nrows=15, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True)

    for df in [north_all_2015, north_all_2016, north_all_2017, north_all_2018, north_all_2019, north_all_2020, north_all_2021, north_all_2022, north_all_2023, north_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].replace(rayon_map_north)

    for df in [north_non_2016, north_non_2017, north_non_2018, north_non_2019, north_non_2020, north_non_2021, north_non_2022, north_non_2023, north_non_2024]:
        df.columns = NON_COLS
        df[['all_nonextr_inv', 'f_nonextr_inv']] = df[['all_nonextr_inv', 'f_nonextr_inv']].astype(str).apply(lambda x: x.str.replace(' ', '')).astype(float)
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].replace(rayon_map_north)

    all_by_year = {
        2015: north_all_2015, 2016: north_all_2016, 2017: north_all_2017, 2018: north_all_2018,
        2019: north_all_2019, 2020: north_all_2020, 2021: north_all_2021, 2022: north_all_2022,
        2023: north_all_2023, 2024: north_all_2024,
    }
    non_by_year = {
        2016: north_non_2016, 2017: north_non_2017, 2018: north_non_2018,
        2019: north_non_2019, 2020: north_non_2020, 2021: north_non_2021, 2022: north_non_2022,
        2023: north_non_2023, 2024: north_non_2024,
    }

    north_year_dfs = []
    for yr in sorted(all_by_year.keys()):
        all_df = all_by_year[yr]
        if yr in non_by_year:
            df = all_df.merge(non_by_year[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'North-Kazakhstan'
        df['year'] = yr
        north_year_dfs.append(df)

    north_inv = pd.concat(north_year_dfs, ignore_index=True)
    north_inv['rayon'] = north_inv['rayon'].str.strip()

    # --- Employment ---
    north_empl = pd.read_excel(Path("../Data/Employment/North-Kazakhstan.xlsx"), skiprows=27, nrows=14, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    north_empl.columns = ['rayon'] + list(range(2014, 2025))
    north_empl['rayon'] = north_empl['rayon'].str.strip().replace(rayon_map_north)
    north_employment = north_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/North-Kazakhstan_All.xlsx")
    ctrl_years = list(range(2014, 2025))
    sheet_map = {
        'Petropavl city': ('г.Петропавловск', 'B, Y:AI', [85, 113, 156]),
        'Aiyrtau': ('Айыртауский', 'B, Y:AI', [85, 113, 156]),
        'Aqzhar': ('Акжарский', 'A, Y:AI', [85, 113, 156]),
        'Magzhan Zhumabayev': ('М.Жумабаева', 'A, Y:AI', [85, 113, 156]),
        'Esil': ('Есильский', 'A, Y:AI', [85, 113, 156]),
        'Zhambyl': ('Жамбылский', 'A, Y:AI', [85, 113, 156]),
        'Qyzylzhar': ('Кызылжарский', 'A, Y:AI', [85, 113, 156]),
        'Mamlyut': ('Мамлютский', 'A, Y:AI', [85, 113, 156]),
        'Shal Aqyn': ('Шал акына', 'A, Y:AI', [85, 113, 156]),
        'Aqqayin': ('Аккайынский', 'A, Y:AI', [85, 113, 156]),
        'Taiynsha': ('Тайыншинский', 'A, Y:AI', [85, 113, 156]),
        'Timiryazev': ('Тимирязевский', 'A, Y:AI', [85, 113, 156]),
        'Ualikhanov': ('Уалихановский', 'A, Y:AI', [85, 113, 156]),
        'Gabit Musrepov': ('Г.Мусрепова', 'A, Y:AI', [85, 113, 156]),
    }
    north_ctrl_dfs = []
    for rayon_name, (sheet, ucols, rows) in sheet_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, ucols, 180, ctrl_years, rows, rayon_name)
        north_ctrl_dfs.append(df)

    north_controls = pd.concat(north_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/North-Kazakhstan.xlsx"), header=2, usecols='Y:AI', nrows=199, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(north_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    north_ppi = pd.DataFrame(ppi_records)
    north_ppi['ppi'] = clean_col(north_ppi['ppi'])

    north_merged = outer_merge([north_inv, north_employment, north_controls, north_ppi])
    north_merged['oblast'] = north_merged['oblast'].fillna('North-Kazakhstan')
    return north_merged


def process_pavlodar():
    base = Path("../Data/Investments/Pavlodar")

    rayon_map_pavlodar = {
        'Павлодар қ.ə.': 'Pavlodar city', 'Павлодар қ.ә.': 'Pavlodar city', 'Павлодар қ.ә': 'Pavlodar city',
        'Павлодар\xa0қ.ә.': 'Pavlodar city', 'Павлодар\xa0қ.ә': 'Pavlodar city',
        'Павлодар г.а.': 'Pavlodar city', 'Павлодар  г.а.': 'Pavlodar city', 'Павлодар ': 'Pavlodar',
        'Ақсу қ.ə.': 'Aqsu city', 'Ақсу қ.ә.': 'Aqsu city', 'Ақсу қ.ә': 'Aqsu city',
        'Ақсу\xa0қ.ә.': 'Aqsu city', 'Ақсу\xa0қ.ә': 'Aqsu city', 'Аксу  г.а.': 'Aqsu city', 'Аксу г.а.': 'Aqsu city',
        'Екібастұз қ.ə.': 'Ekibastuz city', 'Екібастұз қ.ә.': 'Ekibastuz city', 'Екібастұз қ.ә': 'Ekibastuz city',
        'Екібастұз\xa0қ.ә.': 'Ekibastuz city', 'Екібастұз\xa0қ.ә': 'Ekibastuz city', 'Екибастуз  г.а.': 'Ekibastuz city', 'Экибастуз  г.а.': 'Ekibastuz city', 'Экибастуз г.а.': 'Ekibastuz city',
        'Ақтоғай': 'Aqtoghay', 'Ақтоғай ': 'Aqtoghay', 'Актогайский ': 'Aqtoghay', 'Актогайский': 'Aqtoghay',
        'Баянауыл': 'Bayanauyl', 'Баянауыл ': 'Bayanauyl', 'Баянаульский ': 'Bayanauyl', 'Баянаульский': 'Bayanauyl',
        'Железин': 'Zhelezin', 'Железин ': 'Zhelezin', 'Железинский ': 'Zhelezin', 'Железинский': 'Zhelezin',
        'Ертіс': 'Ertis', 'Ертіс ': 'Ertis', 'Иртышский': 'Ertis', 'Иртышский ': 'Ertis',
        'Качиры': 'Terengkol', 'Лебяжі': 'Aqquly', 'Павлодар': 'Pavlodar',
        'Май': 'May', 'Май ': 'May', 'Майский ': 'May', 'Майский': 'May',
        'Успен': 'Uspen', 'Успен ': 'Uspen', 'Успенский': 'Uspen', 'Успенский ': 'Uspen',
        'Шарбақты': 'Sharbaqty', 'Шарбақты ': 'Sharbaqty', 'Шарбақты ауданы': 'Sharbaqty', 'Щербактинский': 'Sharbaqty', 'Щербактинский ': 'Sharbaqty',
        'Тереңкөл': 'Terengkol', 'Аққулы': 'Aqquly',
        'Павлодарский': 'Pavlodar', 'Павлодарский ': 'Pavlodar',
    }

    def read_pav_pdf_all(year, pages, row_start, col_indices, drop_idx=None):
        t = camelot.read_pdf(base / f"{year}_Investments.pdf", pages=pages, flavor='stream')[0].df
        df = t.iloc[row_start:, col_indices].reset_index(drop=True)
        if drop_idx is not None:
            df = df.drop(index=drop_idx).reset_index(drop=True)
        return df

    # 2014-2018 PDF
    pav_all_2014_raw = camelot.read_pdf(base / "2014_Investments.pdf", pages='29', flavor='stream')[0].df
    pav_all_2014 = pav_all_2014_raw.iloc[6:, [0, 2, 5]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2014[5] = pav_all_2014[5].astype(str).replace('–', '')
    for col in [pav_all_2014.columns[1], pav_all_2014.columns[2]]:
        pav_all_2014[col] = pav_all_2014[col].astype(str).str.replace(' ', '', regex=False).replace('', pd.NA).astype('Int64')
    pav_all_2014 = pav_all_2014.drop(index=13).reset_index(drop=True)

    pav_all_2015_raw = camelot.read_pdf(base / "2015_Investments.pdf", pages='22', flavor='stream')[0].df
    pav_all_2015 = pav_all_2015_raw.iloc[7:, [0, 2, 5]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2015[5] = pav_all_2015[5].astype(str).replace('–', '')
    for col in [pav_all_2015.columns[1], pav_all_2015.columns[2]]:
        pav_all_2015[col] = pav_all_2015[col].astype(str).str.replace(' ', '', regex=False).replace('', pd.NA).astype('Int64')
    pav_all_2015 = pav_all_2015.drop(index=13).reset_index(drop=True)

    pav_all_2016_raw = camelot.read_pdf(base / "2016_Investments.pdf", pages='26', flavor='stream')[0].df
    pav_all_2016 = pav_all_2016_raw.iloc[7:, [0, 2, 4]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2016[4] = pav_all_2016[4].astype(str).replace('–', '')
    for col in [pav_all_2016.columns[1], pav_all_2016.columns[2]]:
        pav_all_2016[col] = pav_all_2016[col].astype(str).str.replace('\xa0', '', regex=False).str.replace(' ', '', regex=False).replace('', pd.NA).astype('Int64')
    pav_all_2016 = pav_all_2016.drop(index=13).reset_index(drop=True)

    pav_all_2017_raw = camelot.read_pdf(base / "2017_Investments.pdf", pages='30', flavor='stream')[0].df
    pav_all_2017 = pav_all_2017_raw.iloc[7:, [0, 2, 4]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    for col in [pav_all_2017.columns[1], pav_all_2017.columns[2]]:
        pav_all_2017[col] = pav_all_2017[col].astype(str).str.replace(' ', '', regex=False).str.replace('\xa0', '', regex=False).astype(int)

    pav_all_2018_raw = camelot.read_pdf(base / "2018_Investments.pdf", pages='27', flavor='stream')[0].df
    pav_all_2018 = pav_all_2018_raw.iloc[7:, [0, 3, 6]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2018[6] = pav_all_2018[6].astype(str).replace('‐', '')
    for col in [pav_all_2018.columns[1], pav_all_2018.columns[2]]:
        pav_all_2018[col] = pav_all_2018[col].astype(str).str.replace('\xa0', '', regex=False).replace('', pd.NA).astype('Int64')

    pav_all_2019 = pd.read_excel(base / "2019_Investments.xls", sheet_name='14.', header=4, nrows=16, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2020 = pd.read_excel(base / "2020_Investments.xls", sheet_name='13.', header=3, nrows=16, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2021 = pd.read_excel(base / "2021_Investments.xls", sheet_name='13.', header=3, nrows=16, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2022 = pd.read_excel(base / "2022_Investments.xlsx", sheet_name='21', header=3, nrows=16, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2023 = pd.read_excel(base / "2023_Investments.xlsx", sheet_name='21', header=3, nrows=16, usecols='A, C, F', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_all_2024 = pd.read_excel(base / "2024_Investments.xlsx", sheet_name='14', header=3, nrows=16, usecols='A, C, E', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)

    # non
    pav_non_2015_raw = camelot.read_pdf(base / "2015_Investments.pdf", pages='23', flavor='stream')[0].df
    pav_non_2015 = pav_non_2015_raw.iloc[5:, [0, 2, 4]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2015[4] = pav_non_2015[4].astype(str).replace('–', '')
    for col in [pav_non_2015.columns[1], pav_non_2015.columns[2]]:
        pav_non_2015[col] = pav_non_2015[col].astype(str).str.replace(' ', '', regex=False).replace('', pd.NA).astype('Int64')
    pav_non_2015 = pav_non_2015.drop(index=13).reset_index(drop=True)

    pav_non_2016_raw = camelot.read_pdf(base / "2016_Investments.pdf", pages='27', flavor='stream')[0].df
    pav_non_2016 = pav_non_2016_raw.iloc[4:, [0, 2, 6]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2016[6] = pav_non_2016[6].astype(str).replace('–', '')
    for col in [pav_non_2016.columns[1], pav_non_2016.columns[2]]:
        pav_non_2016[col] = pav_non_2016[col].astype(str).str.replace('\xa0', '', regex=False).replace('', pd.NA).astype('Int64')
    pav_non_2016 = pav_non_2016.drop(index=13).reset_index(drop=True)

    pav_non_2017_raw = camelot.read_pdf(base / "2017_Investments.pdf", pages='31', flavor='stream')[0].df
    pav_non_2017 = pav_non_2017_raw.iloc[5:, [0, 2, 5]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2017[5] = pav_non_2017[5].astype(str).replace('‐', '')
    for col in [pav_non_2017.columns[1], pav_non_2017.columns[2]]:
        pav_non_2017[col] = pav_non_2017[col].astype(str).str.replace('\xa0', '', regex=False).replace('', pd.NA).astype('Int64')

    pav_non_2018_raw = camelot.read_pdf(base / "2018_Investments.pdf", pages='28', flavor='stream')[0].df
    pav_non_2018 = pav_non_2018_raw.iloc[4:, [0, 2, 5]].reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2018[5] = pav_non_2018[5].astype(str).replace('‐', '')
    for col in [pav_non_2018.columns[1], pav_non_2018.columns[2]]:
        pav_non_2018[col] = pav_non_2018[col].astype(str).str.replace('\xa0', '', regex=False).replace('', pd.NA).astype('Int64')

    pav_non_2019 = pd.read_excel(base / "2019_Investments.xls", sheet_name='15.', header=5, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2020 = pd.read_excel(base / "2020_Investments.xls", sheet_name='14.', header=3, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2021 = pd.read_excel(base / "2021_Investments.xls", sheet_name='14.', header=3, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2022 = pd.read_excel(base / "2022_Investments.xlsx", sheet_name='22', header=3, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2023 = pd.read_excel(base / "2023_Investments.xlsx", sheet_name='22', header=3, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_non_2024 = pd.read_excel(base / "2024_Investments.xlsx", sheet_name='15', header=3, nrows=16, usecols='A, B, D', na_values=NA_VALS).fillna(0).drop(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)

    for df in [pav_all_2014, pav_all_2015, pav_all_2016, pav_all_2017, pav_all_2018, pav_all_2019, pav_all_2020, pav_all_2021, pav_all_2022, pav_all_2023, pav_all_2024]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].replace(rayon_map_pavlodar)

    for df in [pav_non_2015, pav_non_2016, pav_non_2017, pav_non_2018, pav_non_2019, pav_non_2020, pav_non_2021, pav_non_2022, pav_non_2023, pav_non_2024]:
        df.columns = NON_COLS
        # Convert to numeric robustly: dashes in the source PDFs were mapped to <NA>
        # upstream; coerce anything non-numeric to NaN, then fill with 0.
        # (The previous version string-matched the literal '<NA>', which broke
        # silently under pandas >= 3.0 where NA no longer stringifies to '<NA>'.)
        df[['all_nonextr_inv', 'f_nonextr_inv']] = (
            df[['all_nonextr_inv', 'f_nonextr_inv']]
            .apply(lambda s: pd.to_numeric(
                s.astype(str).str.replace(' ', '', regex=False).str.replace('\xa0', '', regex=False),
                errors='coerce'))
            .fillna(0)
        )
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].replace(rayon_map_pavlodar)

    all_by_year = {
        2014: pav_all_2014, 2015: pav_all_2015, 2016: pav_all_2016, 2017: pav_all_2017, 2018: pav_all_2018,
        2019: pav_all_2019, 2020: pav_all_2020, 2021: pav_all_2021, 2022: pav_all_2022, 2023: pav_all_2023, 2024: pav_all_2024,
    }
    non_by_year = {
        2015: pav_non_2015, 2016: pav_non_2016, 2017: pav_non_2017, 2018: pav_non_2018,
        2019: pav_non_2019, 2020: pav_non_2020, 2021: pav_non_2021, 2022: pav_non_2022, 2023: pav_non_2023, 2024: pav_non_2024,
    }

    pav_year_dfs = []
    for yr in sorted(all_by_year.keys()):
        all_df = all_by_year[yr]
        if yr in non_by_year:
            df = all_df.merge(non_by_year[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Pavlodar'
        df['year'] = yr
        pav_year_dfs.append(df)

    pavlodar_inv = pd.concat(pav_year_dfs, ignore_index=True)
    pavlodar_inv['rayon'] = pavlodar_inv['rayon'].str.strip()

    # --- Employment ---
    pav_empl = pd.read_excel(Path("../Data/Employment/Pavlodar.xlsx"), skiprows=25, nrows=14, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True).drop(index=3).reset_index(drop=True)
    pav_empl.columns = ['rayon'] + list(range(2014, 2025))
    pav_empl['rayon'] = pav_empl['rayon'].str.strip().replace(rayon_map_pavlodar)
    pavlodar_employment = pav_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Pavlodar_All.xlsx")
    ctrl_years = list(range(2014, 2025))
    sheet_rows_map = {
        'Pavlodar city': ('Павлодар г.а. ', [89, 118, 163]),
        'Aqsu city': ('Аксу г.а. ', [86, 112, 155]),
        'Ekibastuz city': ('Экибастуз г.а.', [88, 116, 160]),
        'Aqtoghay': ('Актогайский район', [65, 107, 151]),
        'Bayanauyl': ('Баянаульский район', [86, 108, 152]),
        'Zhelezin': ('Железинский район', [86, 107, 151]),
        'Ertis': ('Иртышский район', [86, 106, 150]),
        'Terengkol': ('Теренкөл район', [87, 108, 152]),
        'Aqquly': ('Аққулы район', [86, 108, 151]),
        'May': ('Майский район', [86, 108, 152]),
        'Pavlodar': ('Павлодарский район', [85, 109, 152]),
        'Uspen': ('Успенский район', [87, 109, 153]),
        'Sharbaqty': ('Щербактинский район', [86, 109, 152]),
    }
    pav_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AI', 180, ctrl_years, rows, rayon_name)
        pav_ctrl_dfs.append(df)

    pavlodar_controls = pd.concat(pav_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Pavlodar.xlsx"), header=2, usecols='Y:AI', nrows=83, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(pavlodar_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    pavlodar_ppi = pd.DataFrame(ppi_records)
    pavlodar_ppi['ppi'] = clean_col(pavlodar_ppi['ppi'])

    pavlodar_merged = outer_merge([pavlodar_inv, pavlodar_employment, pavlodar_controls, pavlodar_ppi])
    pavlodar_merged['oblast'] = pavlodar_merged['oblast'].fillna('Pavlodar')
    return pavlodar_merged


def process_shymkent_city():
    base = Path("../Data/Investments/Shymkent-city")

    rayon_map_shymkent = {
        'Шымкент қаласы': 'Shymkent city', 'Шымкент қ.': 'Shymkent city',
        'г. Шымкент': 'Shymkent city', 'г.Шымкент': 'Shymkent city',
        'город Шымкент': 'Shymkent city',
        'Абай ауданы': 'Abay', 'Абайский район': 'Abay',
        'Әл  Фараби': 'Al-Farabi', 'Әл  Фараби ': 'Al-Farabi',
        'Әл-Фараби': 'Al-Farabi', 'Әл-Фараби ': 'Al-Farabi',
        'Әл -Фараби ': 'Al-Farabi', 'Әл-Фараби ауданы': 'Al-Farabi',
        'Аль-Фарабийский': 'Al-Farabi', 'Аль-Фарабийский ': 'Al-Farabi',
        'Аль-Фарабийский район': 'Al-Farabi',
        'Еңбекші': 'Enbekshi', 'Еңбекші ': 'Enbekshi', 'Еңбекші ауданы': 'Enbekshi',
        'Енбекшинский': 'Enbekshi', 'Енбекшинский ': 'Enbekshi',
        'Енбекшинский район': 'Enbekshi',
        'Қаратау': 'Qaratau', 'Қаратау ': 'Qaratau', 'Қаратау ауданы': 'Qaratau',
        'Каратауский': 'Qaratau', 'Каратауский ': 'Qaratau',
        'Каратауский район': 'Qaratau',
        'Тұран': 'Turan', 'Тұран ауданы': 'Turan', 'Туранский район': 'Turan',
    }

    def read_shym_all(path, sheet, header, nrows, usecols):
        kw = dict(header=header, nrows=nrows, usecols=usecols, na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_shymkent)
        return df

    def read_shym_non(path, sheet, header, nrows):
        kw = dict(header=header, nrows=nrows, usecols='A, B, D', na_values=NA_VALS)
        if sheet:
            kw['sheet_name'] = sheet
        df = pd.read_excel(path, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_shymkent)
        return df

    all_2020 = read_shym_all(base / "2020_Investments.xls", '12', 3, 6, 'A, C, E')
    all_2021 = read_shym_all(base / "2021_Investments.xls", '12', 3, 15, 'A, C, E')
    all_2022 = read_shym_all(base / "2022_Investments.xlsx", '20', 2, 6, 'A, C, E')
    all_2023 = read_shym_all(base / "2023_Investments.xlsx", '20', 2, 8, 'A, C, E')
    all_2024 = read_shym_all(base / "2024_Investments.xlsx", '12', 2, 7, 'A, C, F')

    non_2020 = read_shym_non(base / "2020_Investments.xls", '13', 3, 6)
    non_2021 = read_shym_non(base / "2021_Investments.xls", '13', 3, 6)
    non_2022 = read_shym_non(base / "2022_Investments.xlsx", '21', 2, 6)
    non_2023 = read_shym_non(base / "2023_Investments.xlsx", '21', 2, 7)
    non_2024 = read_shym_non(base / "2024_Investments.xlsx", '13', 2, 7)

    all_by_year = {2020: all_2020, 2021: all_2021, 2022: all_2022, 2023: all_2023, 2024: all_2024}
    non_by_year = {2020: non_2020, 2021: non_2021, 2022: non_2022, 2023: non_2023, 2024: non_2024}

    shymkent_year_dfs = []
    for yr in sorted(all_by_year.keys()):
        df = all_by_year[yr].merge(non_by_year[yr], on='rayon', how='left')
        df['oblast'] = 'Shymkent-city'
        df['year'] = yr
        shymkent_year_dfs.append(df)

    shymkent_city_inv = pd.concat(shymkent_year_dfs, ignore_index=True)
    shymkent_city_inv['rayon'] = shymkent_city_inv['rayon'].str.strip()
    shymkent_city_inv = shymkent_city_inv[shymkent_city_inv['rayon'] != 'Shymkent city'].reset_index(drop=True)

    # --- Employment ---
    shym_empl = pd.read_excel(Path("../Data/Employment/Shymkent-city.xlsx"), skiprows=15, nrows=4, header=None, usecols='A:F, R', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    shym_empl.columns = ['rayon', 2018, 2019, 2020, 2021, 2022, 2023]
    shym_empl['rayon'] = shym_empl['rayon'].str.strip().replace(rayon_map_shymkent)
    shymkent_employment = shym_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Shymkent-city_All.xls")
    ctrl_years = list(range(2017, 2024))
    sheet_rows_map = {
        'Abay': ('Абайский район', [65, 79, 106]),
        'Al-Farabi': ('Аль-Фарабийский район', [65, 79, 106]),
        'Enbekshi': ('Енбекшинский район', [65, 79, 106]),
        'Qaratau': ('Каратауский район', [65, 79, 106]),
    }
    shymkent_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'B:I', 180, ctrl_years, rows, rayon_name)
        shymkent_ctrl_dfs.append(df)

    # Turan: only 1 year (2023)
    turan_df = ctrl_sheet_to_long(ctrl_path, 'Тұран', 2, 'B:C', 180, [2023], [65, 79, 106], 'Turan')
    shymkent_ctrl_dfs.append(turan_df)

    shymkent_controls = pd.concat(shymkent_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Shymkent-city.xls"), header=2, usecols='Y:AH', nrows=54, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(shymkent_city_inv['rayon'].unique())
    year_list = list(range(2014, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    shymkent_city_ppi = pd.DataFrame(ppi_records)
    shymkent_city_ppi['ppi'] = clean_col(shymkent_city_ppi['ppi'])
    # pre-2018 PPI not meaningful for Shymkent city (created 2018)
    shymkent_city_ppi.loc[shymkent_city_ppi['year'].isin([2014, 2015, 2016, 2017]), 'ppi'] = float('nan')

    shymkent_city_merged = outer_merge([shymkent_city_inv, shymkent_employment, shymkent_controls, shymkent_city_ppi])
    shymkent_city_merged['oblast'] = shymkent_city_merged['oblast'].fillna('Shymkent-city')
    return shymkent_city_merged


def process_turkestan():
    base = Path("../Data/Investments/Turkistan")

    rayon_map_turkestan = {
        'Шымкент қ.': 'Shymkent city', 'г.Шымкент': 'Shymkent city',
        'г. Шымкент': 'Shymkent city', 'г.Шымкент ': 'Shymkent city',
        'Арыс қ.': 'Arys city', 'Арыс қ.ə.': 'Arys city', 'Арыс қ.ә.': 'Arys city',
        'г.а.Арысь': 'Arys city', 'г.а. Арыс': 'Arys city', 'Арыс г.а.': 'Arys city',
        'Кентау қ.': 'Kentau city', 'Кентау қ.ə.': 'Kentau city', 'Кентау қ.ә.': 'Kentau city',
        'г.а.Кентау': 'Kentau city', 'г.а. Кентау': 'Kentau city', 'Кентау г.а.': 'Kentau city',
        'Түркістан қ.': 'Turkistan city', 'Түркістан қ.ə.': 'Turkistan city',
        'Түркістан қ.ә.': 'Turkistan city',
        'г.Туркестан': 'Turkistan city', 'г.а. Туркестан': 'Turkistan city',
        'Туркестан г.а.': 'Turkistan city',
        'Бəйдібек ауданы': 'Baidibek', 'Бәйдібек ауданы': 'Baidibek',
        'Байдібек ауданы': 'Baidibek', 'Байдибекский район': 'Baidibek',
        'район Байдибека': 'Baidibek', 'Район Байдибека': 'Baidibek',
        'Қазыгүрт ауданы': 'Qazygurt', 'Қазығұрт': 'Qazygurt', 'Қазығұрт ауданы': 'Qazygurt',
        'Казыгуртский': 'Qazygurt', 'Казыгуртский район': 'Qazygurt',
        'Мақтаарал ауданы': 'Maqtaaral', 'Мақтаарал': 'Maqtaaral',
        'Махтааральский': 'Maqtaaral', 'Мактааральский район': 'Maqtaaral',
        'Ордабасы ауданы': 'Ordabasy', 'Ордабасы': 'Ordabasy',
        'Ордабасинский': 'Ordabasy', 'Ордабасынский': 'Ordabasy',
        'Ордабасынский район': 'Ordabasy',
        'Отырар ауданы': 'Otyrar', 'Отырар': 'Otyrar',
        'Отырарский': 'Otyrar', 'Отрарский': 'Otyrar', 'Отрарский район': 'Otyrar',
        'Сайрам ауданы': 'Sairam', 'Сайрам': 'Sairam',
        'Сайрамский': 'Sairam', 'Сайрамский район': 'Sairam',
        'Сарыағаш ауданы': 'Saryagash', 'Сарыағаш': 'Saryagash',
        'Сарыагашский': 'Saryagash', 'Сарыагашский район': 'Saryagash',
        'Созақ ауданы': 'Sozaq', 'Созақ': 'Sozaq',
        'Сузакский': 'Sozaq', 'Сузакский район': 'Sozaq',
        'Төлеби ауданы': 'Tolebi', 'Төле би ауданы': 'Tolebi', 'Төлеби': 'Tolebi',
        'Толебийский': 'Tolebi', 'Толебийский район': 'Tolebi',
        'Түлкібас ауданы': 'Tulkibas', 'Түлкібас': 'Tulkibas',
        'Тулькубасский': 'Tulkibas', 'Тюлькубасский район': 'Tulkibas',
        'Шардара ауданы': 'Shardara', 'Шардара': 'Shardara',
        'Шардаринский': 'Shardara', 'Шардаринский район': 'Shardara',
        'Жетісай ауданы': 'Zhetysay',
        'Жетысайский': 'Zhetysay', 'Жетысайский район': 'Zhetysay',
        'Келес ауданы': 'Keles',
        'Келесский': 'Keles', 'Келесский район': 'Keles',
        'Сауран ауданы': 'Sauran',
        'Сауранский': 'Sauran', 'Район Сауран': 'Sauran',
    }

    # 2014-2015 PDF via camelot
    t2014 = camelot.read_pdf(base / "2014_Investments.pdf", pages='44', flavor='stream')[0].df
    turk_all_2014 = t2014.iloc[6:, [0, 2, 5]].reset_index(drop=True)
    for col in [turk_all_2014.columns[1], turk_all_2014.columns[2]]:
        turk_all_2014[col] = turk_all_2014[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    t2015 = camelot.read_pdf(base / "2015_Investments.pdf", pages='54', flavor='stream')[0].df
    turk_all_2015 = t2015.iloc[6:, [0, 2, 4]].reset_index(drop=True)
    for col in [turk_all_2015.columns[1], turk_all_2015.columns[2]]:
        turk_all_2015[col] = turk_all_2015[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    tnon2014 = camelot.read_pdf(base / "2014_Investments.pdf", pages='45', flavor='stream')[0].df
    turk_non_2014 = tnon2014.iloc[8:, [0, 2, 6]].reset_index(drop=True)
    for col in [turk_non_2014.columns[1], turk_non_2014.columns[2]]:
        turk_non_2014[col] = turk_non_2014[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    tnon2015 = camelot.read_pdf(base / "2015_Investments.pdf", pages='55', flavor='stream', table_areas=['0,0,1000,1000'])[0].df
    turk_non_2015 = tnon2015.iloc[27:, [0, 2, 6]].reset_index(drop=True).drop(15).reset_index(drop=True)
    for col in [turk_non_2015.columns[1], turk_non_2015.columns[2]]:
        turk_non_2015[col] = turk_non_2015[col].astype(str).str.replace(' ', '', regex=False).astype(int)

    # 2016-2024 Excel
    all_specs = [
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 6, 16, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 6, 16, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 6, 17, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '14', 4, 17, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13', 4, 17, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13', 4, 18, 'A, C, E'),
        (2022, base / "2022_Investments.xls", '13', 4, 18, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '19', 3, 18, 'A, C, F'),
        (2024, base / "2024_Investments.xls", '14', 4, 18, 'A, C, E'),
    ]
    non_specs = [
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 7, 16, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 7, 16, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 7, 17, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '15', 4, 17, 'A, B, D'),
        (2020, base / "2020_Investments.xls", '14', 4, 17, 'A, B, D'),
        (2021, base / "2021_Investments.xls", '14', 4, 18, 'A, B, D'),
        (2022, base / "2022_Investments.xls", '14', 4, 18, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '20', 3, 18, 'A, B, D'),
        (2024, base / "2024_Investments.xls", '15', 4, 18, 'A, B, D'),
    ]

    all_dfs = {2014: turk_all_2014, 2015: turk_all_2015}
    non_dfs = {2014: turk_non_2014, 2015: turk_non_2015}

    for yr, ap, asht, ahdr, anrows, aucols in all_specs:
        kw = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw['sheet_name'] = asht
        df = pd.read_excel(ap, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_turkestan)
        all_dfs[yr] = df

    for yr, np_, nsht, nhdr, nnrows, nusecols in non_specs:
        kw = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw['sheet_name'] = nsht
        df = pd.read_excel(np_, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_turkestan)
        non_dfs[yr] = df

    for df in [turk_all_2014, turk_all_2015]:
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_turkestan)
    for df in [turk_non_2014, turk_non_2015]:
        df.columns = NON_COLS
        df[['all_nonextr_inv', 'f_nonextr_inv']] = df[['all_nonextr_inv', 'f_nonextr_inv']].astype(str).apply(lambda x: x.str.replace(' ', '')).astype(float)
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_turkestan)

    turk_year_dfs = []
    for yr in sorted(all_dfs.keys()):
        all_df = all_dfs[yr]
        if yr in non_dfs:
            df = all_df.merge(non_dfs[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Turkistan'
        df['year'] = yr
        turk_year_dfs.append(df)

    turkestan_inv = pd.concat(turk_year_dfs, ignore_index=True)
    turkestan_inv['rayon'] = turkestan_inv['rayon'].str.strip()

    # --- Employment ---
    turk_empl = pd.read_excel(Path("../Data/Employment/Turkistan.xls"), skiprows=28, nrows=18, header=None, usecols='A:K', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    turk_empl.columns = ['rayon'] + list(range(2014, 2024))
    turk_empl['rayon'] = turk_empl['rayon'].str.strip().replace(rayon_map_turkestan)
    turkestan_employment = turk_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Turkistan_All.xls")
    ctrl_years = list(range(2014, 2025))

    # shymkent_city: special params (years 2013-2017, filter to 2014+)
    shymkent_ctrl = ctrl_sheet_to_long(ctrl_path, 'г.Шымкент', 2, 'A, X:AB', 180, list(range(2013, 2018)), [99, 128, 172], 'Shymkent city')
    shymkent_ctrl = shymkent_ctrl[shymkent_ctrl['year'] >= 2014]

    standard_sheets = {
        'Turkistan city': 'г.Туркестан',
        'Arys city': 'г.а.Арысь',
        'Kentau city': 'г.а.Кентау',
        'Baidibek': 'Байдибек',
        'Zhetysay': 'Жетысай',
        'Keles': 'Келес',
        'Qazygurt': 'Казыгурт',
        'Maqtaaral': 'Мактаарал',
        'Ordabasy': 'Ордабасы',
        'Otyrar': 'Отрар',
        'Sairam': 'Сайрам',
        'Saryagash': 'Сарыагаш',
        'Sauran': 'Сауран',
        'Sozaq': 'Сузак',
        'Tolebi': 'Толеби',
        'Tulkibas': 'Тюлькубас',
        'Shardara': 'Шардара',
    }
    turk_ctrl_dfs = [shymkent_ctrl]
    for rayon_name, sheet in standard_sheets.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AI', 180, ctrl_years, [100, 129, 173], rayon_name)
        turk_ctrl_dfs.append(df)

    turkestan_controls = pd.concat(turk_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Turkistan.xls"), header=2, usecols='Y:AI', nrows=64, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(turkestan_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    turkestan_ppi = pd.DataFrame(ppi_records)
    turkestan_ppi['ppi'] = clean_col(turkestan_ppi['ppi'])

    turkestan_merged = outer_merge([turkestan_inv, turkestan_employment, turkestan_controls, turkestan_ppi])
    turkestan_merged['oblast'] = turkestan_merged['oblast'].fillna('Turkistan')
    return turkestan_merged


def process_west_kazakhstan():
    base = Path("../Data/Investments/West-Kazakhstan")

    rayon_map_west = {
        'Орал қ.': 'Oral city', 'Орал қ.ә': 'Oral city', 'Орал қ.ә.': 'Oral city',
        'г.Уральск': 'Oral city', 'Уральск г.а.': 'Oral city', 'г.Уральск ': 'Oral city',
        'Ақжайық': 'Aqzhaiyq', 'Ақжайық ': 'Aqzhaiyq', 'Ақжайық ауданы': 'Aqzhaiyq', 'Акжаикский': 'Aqzhaiyq', 'Акжаик': 'Aqzhaiyq',
        'Бөкей ордасы': 'Bokey Orda', 'Бөкей ордасы  ': 'Bokey Orda',
        'Бөкейордасы  ауданы': 'Bokey Orda', 'Бөкейордасы ауданы': 'Bokey Orda',
        'Бокейординский': 'Bokey Orda', 'Бокей орда': 'Bokey Orda',
        'Бөрлі': 'Borili', 'Бөрлі ': 'Borili', 'Бөрлі ауданы': 'Borili', 'Бурлинский': 'Borili', 'Бурлин': 'Borili',
        'Жаңғала': 'Zhangala', 'Жаңақала': 'Zhangala', 'Жаңақала ': 'Zhangala', 'Жаңақала ауданы': 'Zhangala',
        'Жангалинский': 'Zhangala', 'Жангала': 'Zhangala',
        'Жәнібек': 'Zhanibek', 'Жәнібек ': 'Zhanibek', 'Жәнібек ауданы': 'Zhanibek',
        'Жанибекский': 'Zhanibek', 'Жанибек': 'Zhanibek',
        'Зеленов': 'Baiterek', 'Зеленов ': 'Baiterek', 'Зеленов ауданы': 'Baiterek',
        'Бәйтерек': 'Baiterek', 'Бәйтерек ': 'Baiterek',
        'Казталов': 'Qaztalov', 'Казталов ': 'Qaztalov', 'Казталовка ауданы': 'Qaztalov',
        'Казталовский': 'Qaztalov', 'Казталовский ': 'Qaztalov',
        'Қаратөбе': 'Qaratobe', 'Қаратөбе ': 'Qaratobe', 'Қаратөбе ауданы': 'Qaratobe',
        'Каратобинский': 'Qaratobe', 'Каратобе': 'Qaratobe',
        'Сырым ': 'Syrym', 'Сырым': 'Syrym', 'Сырым ауданы': 'Syrym', 'Сырымский': 'Syrym',
        'Таскала': 'Tasqala', 'Тасқала': 'Tasqala', 'Тасқала ': 'Tasqala',
        'Тасқала ауданы': 'Tasqala', 'Таскалинский': 'Tasqala', 'Таскала ': 'Tasqala',
        'Теректі': 'Terekti', 'Теректі ': 'Terekti', 'Теректі ауданы': 'Terekti',
        'Теректинский': 'Terekti', 'Теректа': 'Terekti',
        'Шыңғырлау': 'Shyngyrlau', 'Шыңғырлау ': 'Shyngyrlau', 'Шыңғырлау ауданы': 'Shyngyrlau',
        'Чингирлауский': 'Shyngyrlau', 'Чингирлау': 'Shyngyrlau',
    }

    all_specs = [
        (2014, base / "2014_Investments.xls", None, 6, 14, 'A, C, D'),
        (2015, base / "2015_Investments/Internal_vs_external.xls", None, 6, 14, 'A, C, D'),
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 5, 14, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 5, 14, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 5, 14, 'A, C, E'),
        (2019, base / "2019_Investments.xls", '14', 7, 14, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '14', 7, 14, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13', 7, 14, 'A, C, E'),
        (2022, base / "2022_Investments.xls", '12.', 5, 14, 'A, C, E'),
        (2023, base / "2023_Investments.xls", '18', 3, 14, 'A, C, F'),
        (2024, base / "2024_Investments.xls", '14.', 5, 14, 'A, C, E'),
    ]
    non_specs = [
        (2015, base / "2015_Investments/Greenfield_investments.xls", None, 7, 14, 'A, B, C'),
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 5, 14, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 5, 14, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 5, 14, 'A, B, D'),
        (2019, base / "2019_Investments.xls", '15', 7, 14, 'A, B, D'),
        (2020, base / "2020_Investments.xls", '15', 7, 14, 'A, B, D'),
        (2021, base / "2021_Investments.xls", '14', 5, 14, 'A, B, D'),
        (2022, base / "2022_Investments.xls", '13.', 5, 14, 'A, B, D'),
        (2023, base / "2023_Investments.xls", '19', 3, 14, 'A, B, D'),
        (2024, base / "2024_Investments.xls", '15.', 5, 14, 'A, B, D'),
    ]

    all_dfs = {}
    for yr, ap, asht, ahdr, anrows, aucols in all_specs:
        kw = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw['sheet_name'] = asht
        df = pd.read_excel(ap, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        # 2017-2018: strip spaces in numeric cols
        if yr in [2017, 2018]:
            for col in [df.columns[1], df.columns[2]]:
                df[col] = df[col].astype(str).str.replace(' ', '', regex=False).astype(float)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_west)
        all_dfs[yr] = df

    non_dfs = {}
    for yr, np_, nsht, nhdr, nnrows, nusecols in non_specs:
        kw = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw['sheet_name'] = nsht
        df = pd.read_excel(np_, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = NON_COLS
        if yr in [2017, 2018]:
            for col in [df.columns[1], df.columns[2]]:
                df[col] = df[col].astype(str).str.replace(' ', '', regex=False).astype(float)
        df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
        df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_west)
        non_dfs[yr] = df

    west_year_dfs = []
    for yr in sorted(all_dfs.keys()):
        all_df = all_dfs[yr]
        if yr in non_dfs:
            df = all_df.merge(non_dfs[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'West-Kazakhstan'
        df['year'] = yr
        west_year_dfs.append(df)

    west_inv = pd.concat(west_year_dfs, ignore_index=True)
    west_inv['rayon'] = west_inv['rayon'].str.strip()

    # --- Employment ---
    west_empl = pd.read_excel(Path("../Data/Employment/West-Kazakhstan.xlsx"), skiprows=25, nrows=13, header=None, usecols='A:K', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    west_empl.columns = ['rayon'] + list(range(2014, 2024))
    west_empl['rayon'] = west_empl['rayon'].str.strip().replace(rayon_map_west)
    west_employment = west_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/West-Kazakhstan_All.xls")
    ctrl_years = list(range(2014, 2024))
    sheet_rows_map = {
        'Oral city': ('г.а Уральск', [68, 86, 130]),
        'Aqzhaiyq': ('Акжаик', [68, 86, 130]),
        'Bokey Orda': ('Бокейорда', [68, 86, 130]),
        'Borili': ('Бурлин', [68, 86, 130]),
        'Zhangala': ('Жанакала', [68, 86, 130]),
        'Zhanibek': ('Жанибек', [68, 86, 130]),
        'Baiterek': ('Бәйтерек', [68, 86, 130]),
        'Terekti': ('Теректа', [68, 86, 130]),
        'Qaztalov': ('Казталов', [68, 86, 130]),
        'Syrym': ('Сырым', [68, 86, 130]),
        'Qaratobe': ('Каратобе', [68, 86, 130]),
        'Tasqala': ('Таскала', [68, 86, 130]),
        'Terekti': ('Теректа', [68, 86, 130]),
        'Shyngyrlau': ('Чингирлау', [68, 86, 130]),
    }
    west_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AH', 180, ctrl_years, rows, rayon_name)
        west_ctrl_dfs.append(df)

    west_controls = pd.concat(west_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/West-Kazakhstan.xls"), header=2, usecols='Y:AH', nrows=86, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2024))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(west_inv['rayon'].unique())
    year_list = list(range(2014, 2024))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    west_ppi = pd.DataFrame(ppi_records)
    west_ppi['ppi'] = clean_col(west_ppi['ppi'])

    west_merged = outer_merge([west_inv, west_employment, west_controls, west_ppi])
    west_merged['oblast'] = west_merged['oblast'].fillna('West-Kazakhstan')
    return west_merged


def process_zhambyl():
    base = Path("../Data/Investments/Zhambyl")

    rayon_map_zhambyl = {
        'Тараз қ.': 'Taraz city', 'Тараз қаласы': 'Taraz city',
        'г. Тараз': 'Taraz city', 'г.Тараз': 'Taraz city',
        'Байзақ': 'Baizaq', 'Байзақ ауданы': 'Baizaq',
        'Байзакский': 'Baizaq', 'Байзакский ': 'Baizaq', 'Байзакский район': 'Baizaq',
        'Жамбыл': 'Zhambyl', 'Жамбыл ауданы': 'Zhambyl',
        'Жамбылский': 'Zhambyl', 'Жамбылский ': 'Zhambyl', 'Жамбылский район': 'Zhambyl',
        'Жуалы': 'Zhualy', 'Жуалы ауданы': 'Zhualy',
        'Жуалы ': 'Zhualy', 'Жуалынский': 'Zhualy', 'Жуалынский ': 'Zhualy',
        'Жуалынский район': 'Zhualy',
        'Қордай': 'Qorday', 'Қордай ауданы': 'Qorday',
        'Қордай ': 'Qorday', 'Кордайский': 'Qorday', 'Кордайский ': 'Qorday',
        'Кордайский район': 'Qorday',
        'Т.Рысқұлов': 'Rysqulov', 'Т.Рысқұлов атындағы аудан': 'Rysqulov',
        'Тұрар Рысқұлов': 'Rysqulov', 'Турара Рыскулова': 'Rysqulov',
        'район Т.Рыскулова': 'Rysqulov',
        'Меркі': 'Merke', 'Меркі ауданы': 'Merke',
        'Меркенский': 'Merke', 'Меркенский ': 'Merke', 'Меркенский район': 'Merke',
        'Мойынқұм': 'Moyynqum', 'Мойынқұм ауданы': 'Moyynqum',
        'Мойынқұм ': 'Moyynqum', 'Мойынкумский': 'Moyynqum',
        'Мойынкумский ': 'Moyynqum', 'Мойынкумский район': 'Moyynqum',
        'Сарысу': 'Sarysu', 'Сарысу ауданы': 'Sarysu',
        'Сарысу ': 'Sarysu', 'Сарысуский': 'Sarysu',
        'Сарысуский ': 'Sarysu', 'Сарысуский район': 'Sarysu',
        'Талас': 'Talas', 'Талас ауданы': 'Talas',
        'Талас ': 'Talas', 'Таласский': 'Talas',
        'Таласский ': 'Talas', 'Таласский район': 'Talas',
        'Шу': 'Shu', 'Шу ауданы': 'Shu',
        'Шу ': 'Shu', 'Шуский': 'Shu',
        'Шуский ': 'Shu', 'Шуский район': 'Shu',
    }

    # 2015 PDF via camelot
    t2015 = camelot.read_pdf(base / "2015_Investments.pdf", pages='13', flavor='stream')[0].df
    zhambyl_all_2015 = t2015.iloc[4:, [0, 2, 5]].reset_index(drop=True)
    for col in [zhambyl_all_2015.columns[1], zhambyl_all_2015.columns[2]]:
        zhambyl_all_2015[col] = zhambyl_all_2015[col].astype(str).str.replace(' ', '', regex=False).astype(int)
    zhambyl_all_2015.columns = INV_COLS
    to_numeric_inv(zhambyl_all_2015, ['l_all_inv', 'f_all_inv'])
    zhambyl_all_2015['rayon'] = zhambyl_all_2015['rayon'].str.strip().replace(rayon_map_zhambyl)

    all_specs = [
        (2016, base / "2016_Investments/Internal_vs_external.xls", None, 6, 14, 'A, C, E'),
        (2017, base / "2017_Investments/Internal_vs_external.xls", None, 6, 14, 'A, C, E'),
        (2018, base / "2018_Investments/Internal_vs_external.xls", None, 6, 14, 'A, C, E'),
        (2019, base / "2019_Investments.xlsx", '14', 6, 12, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13', 6, 12, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '13', 6, 12, 'A, C, E'),
        (2022, base / "2022_Investments.xls", '14', 6, 12, 'A, C, E'),
        (2023, base / "2023_Investments.xlsx", '19', 4, 12, 'A, C, F'),
        (2024, base / "2024_Investments.xlsx", '14', 4, 12, 'A, C, E'),
    ]
    non_specs = [
        (2016, base / "2016_Investments/Greenfield_investments.xls", None, 7, 12, 'A, B, D'),
        (2017, base / "2017_Investments/Greenfield_investments.xls", None, 7, 12, 'A, B, D'),
        (2018, base / "2018_Investments/Greenfield_investments.xls", None, 7, 12, 'A, B, D'),
        (2019, base / "2019_Investments.xlsx", '14', 6, 12, 'A, C, E'),
        (2020, base / "2020_Investments.xls", '13', 6, 12, 'A, C, E'),
        (2021, base / "2021_Investments.xls", '14', 7, 12, 'A, B, D'),
        (2022, base / "2022_Investments.xls", '15', 7, 12, 'A, B, D'),
        (2023, base / "2023_Investments.xlsx", '20', 4, 12, 'A, B, D'),
        (2024, base / "2024_Investments.xlsx", '15', 4, 12, 'A, B, D'),
    ]

    all_dfs = {2015: zhambyl_all_2015}
    non_dfs = {}

    for yr, ap, asht, ahdr, anrows, aucols in all_specs:
        kw = dict(header=ahdr, nrows=anrows, usecols=aucols, na_values=NA_VALS)
        if asht:
            kw['sheet_name'] = asht
        df = pd.read_excel(ap, **kw).fillna(0).drop(0).reset_index(drop=True)
        df.columns = INV_COLS
        to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_zhambyl)
        all_dfs[yr] = df

    for yr, np_, nsht, nhdr, nnrows, nusecols in non_specs:
        kw = dict(header=nhdr, nrows=nnrows, usecols=nusecols, na_values=NA_VALS)
        if nsht:
            kw['sheet_name'] = nsht
        df = pd.read_excel(np_, **kw).fillna(0).drop(0).reset_index(drop=True)
        # 2019-2020 non use same sheet as all (cols are all_inv format)
        if yr in [2019, 2020]:
            df.columns = INV_COLS
            to_numeric_inv(df, ['l_all_inv', 'f_all_inv'])
            df['l_nonextr_inv'] = df['f_all_inv']
            df['f_nonextr_inv'] = df['f_all_inv']
            df = df[['rayon', 'f_nonextr_inv', 'l_nonextr_inv']]
        else:
            df.columns = NON_COLS
            to_numeric_inv(df, ['all_nonextr_inv', 'f_nonextr_inv'])
            df['l_nonextr_inv'] = df['all_nonextr_inv'] - df['f_nonextr_inv']
            df.drop('all_nonextr_inv', axis=1, inplace=True)
        df['rayon'] = df['rayon'].str.strip().replace(rayon_map_zhambyl)
        non_dfs[yr] = df

    zhambyl_year_dfs = []
    for yr in sorted(all_dfs.keys()):
        all_df = all_dfs[yr]
        if yr in non_dfs:
            df = all_df.merge(non_dfs[yr], on='rayon', how='left')
        else:
            df = all_df.copy()
        df['oblast'] = 'Zhambyl'
        df['year'] = yr
        zhambyl_year_dfs.append(df)

    zhambyl_inv = pd.concat(zhambyl_year_dfs, ignore_index=True)
    zhambyl_inv['rayon'] = zhambyl_inv['rayon'].str.strip()

    # --- Employment ---
    zhambyl_empl = pd.read_excel(Path("../Data/Employment/Zhambyl.xlsx"), skiprows=21, nrows=11, header=None, usecols='A:L', na_values=NA_VALS).fillna(0).reset_index(drop=True)
    zhambyl_empl.columns = ['rayon'] + list(range(2014, 2025))
    zhambyl_empl['rayon'] = zhambyl_empl['rayon'].str.strip().replace(rayon_map_zhambyl)
    zhambyl_employment = zhambyl_empl.melt(id_vars='rayon', var_name='year', value_name='employment')

    # --- Controls ---
    ctrl_path = Path("../Data/Controls/Zhambyl_All.xls")
    ctrl_years = list(range(2014, 2025))
    sheet_rows_map = {
        'Taraz city': ('г.Тараз', [89, 112, 158]),
        'Baizaq': ('Байзакский район', [71, 90, 136]),
        'Zhambyl': ('Жамбылский район', [71, 90, 136]),
        'Zhualy': ('Жуалынский район', [71, 90, 136]),
        'Qorday': ('Кордайский район', [71, 91, 137]),
        'Rysqulov': ('район Т.Рыскулова', [71, 90, 136]),
        'Merke': ('Меркенский район', [71, 91, 137]),
        'Moyynqum': ('Мойынкумский район', [71, 89, 135]),
        'Sarysu': ('Сарысуский район', [85, 104, 150]),
        'Talas': ('Таласский район', [73, 93, 139]),
        'Shu': ('Шуский район', [71, 91, 137]),
    }
    zhambyl_ctrl_dfs = []
    for rayon_name, (sheet, rows) in sheet_rows_map.items():
        df = ctrl_sheet_to_long(ctrl_path, sheet, 2, 'A, Y:AI', 180, ctrl_years, rows, rayon_name)
        zhambyl_ctrl_dfs.append(df)

    zhambyl_controls = pd.concat(zhambyl_ctrl_dfs, ignore_index=True)

    # --- PPI ---
    ppi_raw = pd.read_excel(Path("../Data/PPI/Zhambyl.xls"), header=2, usecols='Y:AI', nrows=78, na_values=NA_VALS)
    ppi_raw.columns = list(range(2014, 2025))
    ppi_vals = ppi_raw.iloc[-1].values
    rayon_list = list(zhambyl_inv['rayon'].unique())
    year_list = list(range(2014, 2025))
    ppi_records = [{'rayon': r, 'year': yr, 'ppi': v} for r in rayon_list for yr, v in zip(year_list, ppi_vals)]
    zhambyl_ppi = pd.DataFrame(ppi_records)
    zhambyl_ppi['ppi'] = clean_col(zhambyl_ppi['ppi'])

    zhambyl_merged = outer_merge([zhambyl_inv, zhambyl_employment, zhambyl_controls, zhambyl_ppi])
    zhambyl_merged['oblast'] = zhambyl_merged['oblast'].fillna('Zhambyl')
    return zhambyl_merged


def main():
    east_merged = process_east_kazakhstan()
    akmola_merged = process_akmola()
    aktobe_merged = process_aktobe()
    alm_city_merged = process_almaty_city()
    almaty_merged = process_almaty()
    astana_city_merged = process_astana_city()
    atyrau_merged = process_atyrau()
    karagandy_merged = process_karagandy()
    kostanay_merged = process_kostanay()
    kyzylorda_merged = process_kyzylorda()
    mangystau_merged = process_mangystau()
    north_merged = process_north_kazakhstan()
    pavlodar_merged = process_pavlodar()
    shymkent_city_merged = process_shymkent_city()
    turkestan_merged = process_turkestan()
    west_merged = process_west_kazakhstan()
    zhambyl_merged = process_zhambyl()

    all_merged = pd.concat([
        east_merged, akmola_merged, aktobe_merged, alm_city_merged, almaty_merged,
        astana_city_merged, atyrau_merged, karagandy_merged, kostanay_merged,
        kyzylorda_merged, mangystau_merged, north_merged, pavlodar_merged,
        shymkent_city_merged, turkestan_merged, west_merged, zhambyl_merged
    ], ignore_index=True)

    all_merged.to_csv(Path("../Data/clean_data.csv"), index=False)
    return all_merged


def validate(df):
    # 1. Duplicate check
    dups = df[df.duplicated(['oblast', 'rayon', 'year'], keep=False)]
    n_dups = len(dups)
    print(f"Duplicate oblast-rayon-year rows: {n_dups}")
    if n_dups:
        print(dups[['oblast', 'rayon', 'year']].value_counts().head(10))
    assert n_dups == 0, "Found duplicate oblast-rayon-year combinations"

    # 2. Rayons per oblast
    print("\nRayons per oblast:")
    print(df.groupby('oblast')['rayon'].nunique().sort_values().to_string())

    # 3. Disappearing rayons
    print("\nDisappearing rayons (present in some years but not all for their oblast):")
    found_any = False
    for oblast, grp in df.groupby('oblast'):
        all_years = sorted(grp['year'].unique())
        missing = []
        for rayon in grp['rayon'].unique():
            rayon_years = set(grp[grp['rayon'] == rayon]['year'])
            absent = [y for y in all_years if y not in rayon_years]
            if absent:
                missing.append((rayon, absent))
        if missing:
            found_any = True
            print(f"  {oblast}:")
            for rayon, absent in missing:
                print(f"    {rayon!r} missing in years: {absent}")
    if not found_any:
        print("  None — all rayons present in every year of their oblast")

    print("\nAll validation checks passed." if n_dups == 0 and not found_any else "\nWARNING: validation issues found.")


if __name__ == "__main__":
    all_merged = main()
    validate(all_merged)