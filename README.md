# Database for meso-level data on investments in fixed capital

## About

This repository contains code that downloads, cleans, and merges Kazakhstan's 'rayon' (county) level data for investments in fixed capital.

## File overview

| Folder | Subfolder | File | Type | Description |
| ------ | --------- | ---- | ---- | ----------- |
| **Code** |  |  |  | Folder that contains all necessary code files |
|  |  | `data_cleaning.py` | Python | Main file for cleaning and combining the data |
|  |  | `get_employment_data.py` | Python | Code that downloads employment data to `Data` folder |
|  |  | `get_investments_data.py` | Python | Code that downloads investments data to `Data` folder |
|  |  | `get_other_controls.py` | Python | Code that downloads data for controls to `Data` folder |
|  |  | `get_ppi_data.py` | Python | Code that downloads PPI data to `Data` folder |
| **Data** |  |  |  | Folder that contains all downloaded data files |
|  | `Controls` |  |  | Subfolder that contains downloaded data for controls |
|  | `Employment` |  |  | Subfolder that contains downloaded data for employment |
|  | `Investments` |  |  | Subfolder that contains downloaded data for investments |
|  | `PPI` |  |  | Subfolder that contains downloaded data for ppi |
|  |  | `clean_data.csv` | CSV | File that contains clean and combined data |
|  |  | `Employment_sources.xlsx` | Excel | File that contains source links for employment data |
|  |  | `Investments_sources.xlsx` | Excel | File that contains source links for investments data |
|  |  | `Other_controls_sources.xlsx` | Excel | File that contains source links for controls data |
|  |  | `PPI.xlsx` | Excel | File that contains source links for PPI data |
| (root) |  | `Investments_download_report_*.txt` | Text | Log of automated investment file downloads: lists successful and failed downloads by oblast-year |
|  |  | `Employment_download_report_*.txt` | Text | Log of automated employment file downloads: lists successful and failed downloads by oblast |
|  |  | `PPI_download_report_*.txt` | Text | Log of automated PPI file downloads: lists successful and failed downloads by oblast |
|  |  | `Controls_download_report_*.txt` | Text | Log of automated controls file downloads: lists successful and failed downloads by oblast and data type |

## Data download

Each `get_*.py` script queries Stat.gov.kz and saves files to the corresponding subfolder of `Data/`. After each run, a timestamped download report is written to the project root summarising which files succeeded and which failed.

Not all files could be downloaded automatically. The download reports list the failed items by name. Those files were retrieved manually by opening the corresponding URL from the `_sources` Excel files (`Investments_sources.xlsx`, `Employment_sources.xlsx`, `Other_controls_sources.xlsx`, `PPI.xlsx`) and downloading the file directly from the browser.

**Download summary (as of last run):**

| Data type | Successful | Failed | Notes on failures |
| --------- | ---------- | ------ | ----------------- |
| Investments | 155 / 177 | 22 | Scattered across 13 oblasts and various years (see report) |
| Employment | 19 / 20 | 1 | Qaraghandy |
| PPI | 20 / 20 | 0 | — |
| Controls | 19 / 25 | 6 | Aktobe, Aqmola, East-Kazakhstan, Qaraghandy, Mangystau, Ulytau |

## Data coverage

`clean_data.csv` is the final merged panel dataset produced by `data_cleaning.py`. It contains **2,277 observations** across **207 rayons**, **17 oblasts**, and **12 years (2013–2024)**. Not every oblast-rayon-year cell is observed; see the missing data notes below.

### Columns

| Column | Description | Unit |
| ------ | ----------- | ---- |
| `rayon` | Rayon (district/county) name | — |
| `oblast` | Oblast (province) name, using pre-2022 administrative boundaries | — |
| `year` | Calendar year | — |
| `l_all_inv` | Local investments in fixed capital, all sectors | thousands of KZT |
| `f_all_inv` | Foreign investments in fixed capital, all sectors | thousands of KZT |
| `l_nonextr_inv` | Local investments in fixed capital, non-extractive sectors | thousands of KZT |
| `f_nonextr_inv` | Foreign investments in fixed capital, non-extractive sectors | thousands of KZT |
| `employment` | Number of employed persons | thousands of persons |
| `agriculture` | Volume of agricultural products | millions of KZT |
| `production` | Volume of industrial production | millions of KZT |
| `retail` | Retail trade turnover | millions of KZT (see note below) |
| `ppi` | Producer price index | index, previous year = 100 |

> **Note on retail units.** The `retail` column shows an unusually wide value range (0.1 – 51,858,310) compared with `agriculture` and `production`. Source files from different oblasts may use different units (thousands vs. millions of KZT). Verify units against the original source files before using this variable in levels.

### Rayons per oblast

| Oblast | Rayons | Year range |
| ------ | ------: | ---------- |
| Almaty | 19 | 2016–2024 |
| Almaty-city | 8 | 2014–2024 |
| Aqmola | 20 | 2014–2024 |
| Aqtobe | 13 | 2016–2024 |
| Astana-city | 4 | 2014–2024 |
| Atyrau | 8 | 2013–2024 |
| East-Kazakhstan | 19 | 2018–2024 |
| Mangystau | 7 | 2014–2024 |
| North-Kazakhstan | 14 | 2014–2024 |
| Pavlodar | 13 | 2014–2024 |
| Qaraghandy | 18 | 2014–2024 |
| Qostanay | 20 | 2014–2024 |
| Qyzylorda | 9 | 2014–2024 |
| Shymkent-city | 5 | 2014–2024 |
| Turkistan | 18 | 2014–2024 |
| West-Kazakhstan | 13 | 2014–2024 |
| Zhambyl | 11 | 2014–2024 |

> **Administrative note.** Three oblasts were split in 2022 (East-Kazakhstan → East-Kazakhstan + Abay; Almaty → Almaty + Zhetisu; Qaraghandy → Qaraghandy + Ulytau). The dataset recombines post-split rayon-level data back into the pre-split historical oblast structure so the panel is consistent across the full time range.

> **Shymkent-city note.** Shymkent became a city of republican significance in 2018 and is covered in this dataset as a separate oblast with 5 districts for 2020–2024. For 2013–2017 the city appears as a single rayon ('Shymkent city') within Turkistan oblast.

### Missing data by variable

Missing values reflect genuine data gaps in the source publications, not processing errors. All NaNs are preserved as-is. Entries below marked *"all N rayons"* mean every rayon in that oblast is missing for the stated years; all other entries are rayon-specific.

**Total investments** (`l_all_inv`, `f_all_inv`)

| Oblast | Rayon(s) | Missing years |
| ------ | -------- | ------------- |
| Aqmola | all 20 rayons | 2015 |
| Aqmola | Qosshi city | 2014–2020 |
| Aqmola | Shortandy | 2021 |
| Astana-city | Baiqonyr | 2014–2017 |
| Atyrau | all 8 rayons | 2013–2015 |
| North-Kazakhstan | all 14 rayons | 2014 |
| Shymkent-city | Abay, Al-Farabi, Enbekshi, Qaratau | 2014–2019 (was part of Turkistan oblast) |
| Shymkent-city | Turan | 2014–2022 (was part of Turkistan oblast) |
| Turkistan | Keles, Zhetysay | 2014–2017 |
| Turkistan | Sauran | 2014–2020 |
| Turkistan | Shymkent city | 2018–2024 (city became separate oblast in 2018) |
| Zhambyl | all 11 rayons | 2014 |

**Non-extractive investments** (`l_nonextr_inv`, `f_nonextr_inv`)

Same gaps as total investments, plus:

| Oblast | Rayon(s) | Missing years |
| ------ | -------- | ------------- |
| Almaty-city | all 8 rayons | 2014–2017 |
| Mangystau | all 7 rayons | 2014–2015 |
| North-Kazakhstan | all 14 rayons | 2015 |
| Pavlodar | all 13 rayons | 2014 |
| Qaraghandy | all 18 rayons | 2014 |
| Qostanay | Fyodorov | 2017–2018 |
| Qostanay | Rudny city | 2016 |
| Qyzylorda | Zhanaqorgan | 2014 |
| Shymkent-city | Al-Farabi | 2021 |
| Shymkent-city | Turan | 2020–2022 |
| West-Kazakhstan | all 13 rayons | 2014 |
| Zhambyl | all 11 rayons | 2015 |

**Employment**

| Oblast | Rayon(s) | Missing years |
| ------ | -------- | ------------- |
| Almaty | all 19 rayons | 2024 |
| Almaty-city | all 8 rayons | 2014–2015 |
| Aqtobe | all 13 rayons | 2024 |
| Astana-city | all 4 rayons | 2024 |
| Atyrau | all 8 rayons | 2013 |
| Qaraghandy | all 18 rayons | 2024 |
| Shymkent-city | all 5 rayons | 2014–2017, 2024 |
| Shymkent-city | Turan | 2018–2023 |
| Turkistan | all 18 rayons | 2024 |
| West-Kazakhstan | all 13 rayons | 2024 |

The 2024 gap across many oblasts reflects the publication schedule: the 2024 employment yearbook had not yet been released at the time of data collection.

**Agriculture**

| Oblast | Rayon(s) | Missing years |
| ------ | -------- | ------------- |
| Almaty | Alakol, Aqsu, Kerbulaq, Koksu, Panfilov, Qaratal, Sarqan, Taldyqorgan_city, Tekeli_city, Yeskeldi | 2016–2021 |
| Almaty | Balqash, Enbekshiqazaq, Ile, Qarasay, Qonayev_city, Rayimbek, Talgar, Uyghur, Zhambyl | 2024 |
| Almaty-city | all 8 rayons | 2014–2015 |
| Aqmola | all 20 rayons | 2024 |
| Aqmola | Qosshi city | 2014–2020 |
| Aqtobe | all 13 rayons | 2024 |
| Aqtobe | Khromtau | 2016–2023 |
| Astana-city | all 4 rayons | 2024 |
| Astana-city | Baiqonyr | 2014–2017 |
| Atyrau | all 8 rayons | 2024 |
| East-Kazakhstan | all 19 rayons | 2024 |
| Mangystau | all 7 rayons | 2024 |
| Qaraghandy | Abay, Balqash_city, Qarazhal_city, Satpayev_city, Shakhtinsk_city, Temirtau_city | all years (2014–2024) |
| Qaraghandy | Ulytau, Zhanaarka, Zhezkazgan_city | 2024 |
| Qostanay | all 20 rayons | 2024 |
| Qyzylorda | all 9 rayons | 2024 |
| Qyzylorda | Baiqonyr city | 2014–2023 |
| Shymkent-city | all 5 rayons | 2014–2016, 2024 |
| Shymkent-city | Turan | 2017–2023 |
| Turkistan | Keles, Zhetysay | 2014–2017 |
| Turkistan | Sauran | 2014–2019 |
| Turkistan | Shymkent city | 2018–2024 |
| West-Kazakhstan | all 13 rayons | 2024 |

**Production** — same pattern as agriculture with the following differences: Almaty-city missing 2024 (not 2014–2015); East-Kazakhstan partially missing in 2024 (10 of 19 rayons); Qostanay / Zhitiqara missing 2014–2023; Qaraghandy / Qarazhal_city and Satpayev_city also missing 2023.

**Retail** — same pattern as production with the following additions: Almaty-city also missing 2017 and 2018 (six of eight rayons); Aqtobe / Shalkar missing 2020; Pavlodar / Aqquly and Aqsu city missing 2019; several Qaraghandy city rayons also missing 2023.

**PPI**

PPI is an oblast-level series broadcast to all rayons within the oblast, so a gap affects the entire oblast. Missing years:

| Oblast | Missing years |
| ------ | ------------- |
| Almaty | 2023–2024 |
| Aqmola | 2024 |
| Aqtobe | 2024 |
| Atyrau | 2024 |
| East-Kazakhstan | 2022, 2024 |
| Mangystau | 2024 |
| Qostanay | 2024 |
| Qyzylorda | 2024 |
| Shymkent-city | 2014–2017, 2024 |
| West-Kazakhstan | 2024 |
