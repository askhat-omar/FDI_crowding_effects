# Database for meso-level data on investments in fixed capital

## About

This repository contains code that downloads Kazakhstan's 'rayon' (county) level data for investments in fixed capital.

## Usage instructions

 Open the terminal in a desired folder and clone this repository (`Git` is assumed to be already installed on your computer):

```bash
git clone https://github.com/askhat-omar/FDI_crowding_effects.git
```

Change directory to the downloaded folder, install virtual environment and activate it (`Anaconda` is assumed to be alredy installed on your computer):

```bash
cd FDI_crowding_effects
conda create -n fdi_crowding --file conda-spec.txt -y
conda activate fdi_crowding
```

Install necessary packages using `pip`:

```bash
pip install -r requirements.txt
```

After successful installation open Jupyter Lab through terminal:

```bash
jupyter lab
```

Open `data_cleaning.ipynb` in the `Code` folder. If `Jupyter` asks to specify `Python` kernel, choose `fdi_crowding`.

## File overview

| Folder | Subfolder | File | Type | Description |
| ------ | --------- | ---- | ---- | ----------- |
| **Code** |  |  |  | Folder that contains all necessary code files |
|  |  | `data_cleaning.ipynb` | Jupyter Notebook | Main file for cleaning and combining the data |
|  |  | `get_employment_data.py` | Python | Code that downloads employment data to `Data` folder |
|  |  | `get_investments_data.py` | Python | Code that downloads investments data to `Data` folder |
|  |  | `get_other_controls.py` | Python | Code that downloads data for controls to `Data` folder |
|  |  | `get_ppi_data.py` | Python | Code that downloads PPI data to `Data` folder |
| **Data** |  |  |  | Folder that contains all downloaded data files |
|  | `Controls` |  |  | Subfolder that contains downloaded data for controls |
|  | `Employment` |  |  | Subfolder that contains downloaded data for employment |
|  | `Investments` |  |  | Subfolder that contains downloaded data for investments |
|  | `PPI` |  |  | Subfolder that contains downloaded data for ppi |
|  |  | `Employment_sources.xlsx` | Excel | File that contains source links for employment data |
|  |  | `Investments_sources.xlsx` | Excel | File that contains source links for investments data |
|  |  | `Other_controls_sources.xlsx` | Excel | File that contains source links for controls data |
|  |  | `PPI.xlsx` | Excel | File that contains source links for PPI data |