# 2026-AI-competition

Code and aggregated data behind the figures of the manuscript and its SI
Appendix. This is the minimal replication package: it contains the single
plotting script and the final aggregated numbers that each figure is drawn
from. It does not contain raw data, data-collection pipelines, or
intermediate files.

# License and data availability

The source code in this repository is licensed under the MIT License. Aggregate data produced by the authors and included in this repository are licensed under the Creative Commons Attribution 4.0 International (CC BY 4.0) License.

Third-party data used in the study are not redistributed through this repository. Information and links for obtaining these data from the original providers are provided in the documentation.

# Credit

- Survey: Lemeng Liang
- Social media: Yuxin Hou
- Scientific publications: Yuxin Hou, Junming Huang
- PhD degree: Zheqing Ma
- Patent: Zheqing Ma
- Job migration: Milly Yang
- AI models and figures: Junming Huang

# Layout

```text
code/visualize.py                                one script, draws every figure
data/survey/survey.xlsx                          Figure 1
data/social-media/social-media.xlsx              Figure 2, Figure S7
data/phd-degree/phd-degree.xlsx                  Figure 3A, Figures S8-S11
data/job-migration/job-migration.xlsx            Figure 3B-C, Figure S12
data/publication/publication.xlsx                Figure 4A, Figure S13
data/patent/patent.xlsx                          Figures S15-S16
data/commercial-models/commercial-models.xlsx    Figure 5
figures/                                         output (PDF)
```

`visualize.py` reads every data file from `../data/<part>/<part>.xlsx`
relative to its own location and writes PDFs to `../figures/`.

# Usage

```bash
pip install numpy pandas matplotlib openpyxl
python3 code/visualize.py                 # all figures
python3 code/visualize.py survey patent   # only the listed parts
```

Parts: `survey`, `social-media`, `phd-degree`, `job-migration`,
`publication`, `patent`, `commercial-models`. Figures 3 and 4 combine two
parts each and are redrawn whenever either part is selected.

# Data files

Every workbook holds aggregated numbers only. Column names are
self-explanatory; the notes below give the source and the definition of each
table.

## survey/survey.xlsx (Figure 1)

Weighted percentages and means computed from four nationally representative
surveys: Pew Research Center American Trends Panel Wave 152 (2024), General
Social Survey (2024), Chinese General Social Survey (2024), and the China
Family Panel Studies 2025 pilot. Sheets `F1_A` (AI exposure and use),
`F1_B+C` (by educational attainment), `F1_D` (attitude items, percent) and
`F1_E` (comfort items, mean on a 0-10 scale). Survey microdata are available
from the respective survey organisations and are not redistributed here.

## social-media/social-media.xlsx (Figure 2, Figure S7)

Sheet `daily`: one row per sampled observation date (the 1st, 10th and 20th
of each month, March 2024 to March 2025; Weibo substitutes 2024-02-29 and
2024-10-02 for two missing dates). Each value is the daily mean of
user-level mean attitude scores (-2 concerned to +2 excited), unsmoothed;
the script applies a three-observation centred moving average before
plotting.

| Column | Definition |
| --- | --- |
| `weibo-deepseek` | Weibo posts classified by DeepSeek (Figure 2, S7) |
| `twitter-deepseek` | Twitter posts of US users classified by DeepSeek (Figure 2, S7) |
| `twitter-gpt` | The same Twitter posts classified by GPT-5-mini (S7) |

Raw posts were obtained from twitterapi.io and from a Weibo data provider
and cannot be redistributed.

## phd-degree/phd-degree.xlsx (Figure 3A, Figures S8-S11)

Sheet `estimates`: estimated annual number of AI-related STEM doctorates,
2015-2025, for four query calibers (`Abstract, with acronyms` is the
baseline used in Figure 3A; the other three are Figures S9-S11). Count =
official STEM doctoral total x in-database AI share; `upper` and `lower`
bound the years whose official STEM total is extrapolated (China from 2023,
United States in 2025). Sheet `in_database_shares`: share of AI-related
dissertations among STEM dissertations in CNKI (China) and ProQuest (United
States), by caliber (Figure S8). Official totals come from the Educational
Statistics Yearbook of China and the NSF/NCSES Survey of Earned Doctorates.

## job-migration/job-migration.xlsx (Figure 3B-C, Figure S12)

Sheet `odds_ratio`: yearly odds ratio of US-to-China versus China-to-US AI
job transitions, for all AI professionals (`all_positions`) and for those
whose highest recorded seniority is director level or above
(`managerial_positions`). Sheets `moves_all` and `moves_managerial`: yearly
counts of adjacent US-China job moves, LinkedIn job-update counts in China
used as the coverage weight, and the reweighted US-to-China count. Computed
from Revelio Labs LinkedIn records accessed through Wharton Research Data
Services (WRDS); individual-level records cannot be redistributed.

> Wharton Research Data Services (WRDS) was used in accessing the Revelio Labs LinkedIn records. This service and the data available thereon constitute valuable intellectual property and trade secrets of WRDS and/or its third-party suppliers.

## publication/publication.xlsx (Figure 4A, Figure S13)

Sheet `stanford_projected`: projected annual number of AI publications in
computer science for China and the United States, 2015-2024, derived from
the 2026 Stanford AI Index (Figure 1.6.1, global counts; Figure 1.6.6,
country shares including an "Unknown" share) as
global count x country share / (1 - unknown share). The AI Index data are
available from Stanford HAI. Sheet `openalex`: yearly counts of AI-related
works in the OpenAlex snapshot of February 2025, identified by title
keywords and assigned to the first country of the corresponding author
(first author if none), for CN, HK, TW and US; Figure S13 plots CN and US.

## patent/patent.xlsx (Figures S15-S16)

Sheet `lens_families`: AI-related simple patent families in Lens.org by
earliest priority jurisdiction, 2015-2024, all families and families cited
by at least one later patent.

Figure 4B (AI patents granted by CNIPA and USPTO, 2015-2023) is drawn from
the patent-level classification released by Fang et al. (NBER Working Paper
35022), which is not redistributed here. To reproduce the panel, obtain that
data set, aggregate grants by granting office and grant year, and add a
sheet `fang_grants` with columns `year`, `China`, `United States` to
`patent.xlsx`; `visualize.py` draws panel B whenever the sheet is present.

## commercial-models/commercial-models.xlsx (Figure 5)

Sheet `Data`: language models with country, developer, release date,
Artificial Analysis Intelligence Index v4.1.1 score and open/closed weights
status, as of 2026-08-30. Sheets `Sources`, `Method` and `Coverage_gaps`
document the source page of every row and the inclusion rules.

# Not included

- SI Figures S1-S4 (survey distributions by demographic group), S5-S6
  (post-level and like-weighted social media averages) and S14 (DBLP
  robustness check) are not generated by this package.
- The Fang et al. patent counts behind Figure 4B (see above).
- Raw survey, social media, dissertation, LinkedIn and patent records, and
  the collection and classification pipelines that produced the aggregated
  tables above.
