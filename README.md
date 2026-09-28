# DDRC Dose-Response Analysis Tool

Developed by Chloe Larson, Francesca Curreli, Loreto Carvallo-Torres, and J. Fraser Glickman for drug-discovery and high-throughput screening workflows at the Fisher Drug Discovery Resource Center (DDRC), The Rockefeller University.

Copyright © 2026 The Rockefeller University.
## License

This software is licensed under the **Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International License (CC BY-NC-ND 4.0)**.

For licensing inquiries or requests for permissions beyond the terms of this license, please contact the Fisher Drug Discovery Resource Center by opening an Issue in this GitHub repository.

## What does this tool do?

This tool takes raw high-throughput screening (HTS) measurements and a separate compound database, links the two files by plate + well, and produces normalized screening results.

The **DDRC Dose-Response Analysis Tool** is a command-line Python program for analyzing compound dose-response data. It summarizes replicate measurements, fits four-parameter logistic (4PL) dose-response curves, calculates IC50 and other curve metrics, generates individual curve plots and plate maps, and produces an Excel summary workbook.

The tool is designed to be useful with real-world assay datasets rather than requiring a single fixed plate layout or file format.

## What the tool does

For each compound, the program:

1. Reads the assay data from a CSV, Excel, or OpenDocument spreadsheet.
2. Lets the user identify the columns containing plate, well, compound ID, concentration, and raw signal/response.
3. Optionally uses positive and negative assay controls to calculate Percent inhibition and plate-level Z′.
4. Automatically determines the number of replicate measurements at each concentration.
5. Calculates the mean response and SEM for each concentration when replicate data are available.
6. Fits a four-parameter logistic (4PL) dose-response curve.
7. Calculates IC50, chi-square, area under the fitted curve, curve class, fit confidence score, and the DDRC Curve Score.
8. Generates an individual dose-response plot for each successful fit.
9. Produces an Excel workbook containing the results, Z′ values, IC50 summary, and plate maps.

The program is intentionally command-line based. It does not require a graphical user interface.

---

## Supported input files

The program accepts:

- **CSV** (`.csv`)
- **Excel** (`.xlsx` and `.xls`)
- **OpenDocument Spreadsheet** (`.ods`), including files created with LibreOffice/OpenOffice

Place one or more input files in the `input` folder. The program displays the available files and lets you select one.

For Excel and OpenDocument files, the program reads the **first worksheet**.

### Input file selection

Example:

```text
Available input files:

1. assay_results.csv
2. experiment_01.xlsx
3. screening_data.ods

Select input file number:
```

The program does not require the input file to have a particular filename.

---

## Folder structure

```text
DDRC_DRC_Analysis_v2/
├── DRCV7.py
├── README.md
├── requirements.txt
├── input/
├── output/
└── examples/
    └── DDRC_DRCV7_test_input.csv
```

Put your assay data into `input/` and run the program from the main `DDRC_DRC_Analysis_v2` folder.

Results are written automatically to `output/`.

**No file paths need to be edited in the Python script.**

---

## Required data

The program interactively asks the user to identify these fields:

- **Plate** — plate identifier
- **Well** — well position such as A1, B12, or P24
- **Compound ID** — identifier for the test compound
- **Concentration** — compound concentration, preferably in µM or another consistent unit
- **Raw signal / response** — measured assay response or an already-normalized response
- **Batch** — optional batch identifier

Column names do not have to match these names exactly. The user selects the appropriate column by number or by column name.

### Concentrations

Concentrations must be numeric and greater than zero for dose-response fitting. Zero or non-numeric concentrations are not used for the 4PL fit.

The concentration unit is not changed by the program. It is displayed on the plots as `Concentration (µM)`; therefore, datasets should normally use µM concentrations.

---

## Assay controls

Positive and negative controls are **optional**.

If controls are present, the program supports three common arrangements:

1. **Columns** — for example, columns 23 and 24
2. **Rows** — for example, rows A and P
3. **Specific wells** — for example, A1,A2,A3

The user identifies which controls are negative and which are positive.

### Control-based normalization

When positive and negative controls are available, the program calculates Percent inhibition from the raw signal using the plate-specific control means:

```text
Percent inhibition = 100 × (Negative control mean − sample signal)
                    / (Negative control mean − Positive control mean)
```

The resulting values are clipped to the range 0–100 for curve analysis.

The program also calculates a plate-level **Z′ factor** when sufficient positive and negative control measurements are available.

### No controls

If the dataset has no assay controls, the program asks whether the supplied response is:

1. **Already Percent inhibition (0–100)**
2. **Raw signal**

This allows the dose-response fitting portion of the tool to be used even when the normalization was performed elsewhere.

If raw signal is analyzed without controls, the program does **not** invent a Percent inhibition value or Z′ factor.

---

## Positive/negative controls versus reference compound

These are different concepts.

### Positive and negative controls

Assay controls define the response range of the assay and can be used to calculate:

- Percent inhibition
- Z′ factor

### Reference compound

A **reference compound is not an assay control**.

It is a known or previously characterized compound whose dose-response curve can be displayed as a comparison curve on the plots of test compounds.

The reference compound is optional. At the prompt:

```text
Reference compound ID for comparison curve (press Enter to skip):
```

press **Enter** to continue without a reference curve.

If an ID is entered that is not present in the dataset, the program reports a warning and continues without the overlay.

The reference compound is removed from the test-compound results so that it is used for comparison rather than treated as another test compound.

---

## Replicates

The number of replicates does **not** need to be specified by the user.

The program determines the number of measurements automatically for each:

```text
Compound ID × Concentration
```

combination.

For each concentration it calculates:

- Mean response
- SEM when more than one measurement is available
- Number of replicate measurements

Replicate number can vary between concentrations.

For example:

```text
3       = three measurements at every concentration
2-3     = replicate number varies between two and three
1       = single measurements; SEM cannot be estimated
```

The program does not discard an entire curve simply because one concentration has fewer replicates.

---

## Dose-response fitting

The program fits a **four-parameter logistic (4PL)** model to the mean response at each concentration.

The model has four parameters:

- Bottom
- Top
- Log IC50
- Hill slope

At least **four usable positive concentrations** are required for a curve to be fitted.

The Results sheet reports:

- Compound
- Batch
- Fit status
- Replicates
- IC50
- Chi-square
- AUC (area under the fitted curve)
- Curve class
- Fit confidence score
- DDRC Curve Score
- Flags
- Plot

### Fit confidence score

The **Fit confidence score** is a heuristic score used to summarize several aspects of the fitted curve, including response amplitude, normalized RMSE, monotonicity, and the number of concentrations.

It is **not** a statistical confidence interval, p-value, or probability that the fitted curve is correct.

---

## DDRC Curve Score

The **DDRC Curve Score** is the principal curve-quality ranking metric developed for the DDRC tool.

It is a heuristic score from **0–100** based on four components:

| Component | Maximum score |
|---|---:|
| Response span | 30 |
| Plateau quality | 30 |
| Chi-square quality | 25 |
| Hill slope sanity | 15 |
| **Total** | **100** |

The score is intended to help prioritize and review dose-response curves. It should not be interpreted as a statistical confidence measure.

The DDRC Curve Score is deliberately separate from the 4PL fit parameters themselves.

---

## Curve classes and flags

The program assigns a curve class based on amplitude, monotonicity, fit quality, and Hill slope.

Additional flags may identify potential issues such as:

- `weak`
- `poor_fit`
- `noisy`
- `few_points`
- `no_plateau`
- `fit_failed`

A failed fit is reported in the Results sheet rather than silently disappearing from the analysis.

---

## Output workbook

The output is an Excel workbook containing:

### Results

The main results table, including the dose-response metrics and a plot for each successful compound.

The plot is placed in the **Plot** column of the Results sheet.

### Zprime

Plate-level Z′ values when appropriate positive and negative controls are available.

### IC50 Heatmap

A summary of compound IC50 values for successful fits.

### Plate Maps

Plate-level response heatmaps.

---

## Output filename

The output filename is automatically generated from the input filename.

The rule is:

```text
first five characters of input filename + _plot_analysis.xlsx
```

For example:

```text
HTS2026_assay_results.xlsx
```

produces:

```text
HTS20_plot_analysis.xlsx
```

If that filename already exists, the program does not overwrite it. Instead it creates a numbered version such as:

```text
HTS20_plot_analysis_2.xlsx
```

---

## Installation

Python 3 is required.

From the `DDRC_DRC_Analysis_v2` folder, install the required packages with:

```bash
python -m pip install -r requirements.txt
```

The required packages are:

```text
pandas
numpy
matplotlib
scipy
openpyxl
odfpy
xlrd
```

These packages provide support for data handling, numerical calculations, curve fitting, plotting, Excel output, and OpenDocument/legacy Excel input.

---

## Running the program

From the `DDRC_DRC_Analysis_v2` folder:

```bash
python DRCV7.py
```

The program will guide the user through:

1. Selecting the input file
2. Selecting an optional reference compound
3. Mapping the required data columns
4. Selecting the assay-control arrangement, if any
5. Choosing the response type when controls are absent
6. Performing the dose-response analysis

The completed workbook will be placed in the `output` folder.

---

## Example data

An example CSV file is included in the `examples` folder.

The example demonstrates a 384-well-style assay with positive and negative controls arranged in columns. It can be used to confirm that the program is installed and running correctly before analyzing experimental data.

---

## Transparency and modification

The tool is intentionally provided as a Python script rather than as a compiled application. Users can inspect, understand, and modify the calculations.

The core analysis uses standard scientific Python libraries and a four-parameter logistic dose-response model. The DDRC Curve Score is an explicitly defined heuristic ranking system rather than a statistical test.

Users should review the results and underlying dose-response curves rather than relying on any single automated score.
