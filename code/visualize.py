#!/usr/bin/env python3
"""Regenerate the figures of the manuscript and its SI Appendix from the
aggregated data in ../data/<part>/<part>.xlsx.

Usage (from any working directory):

    python3 visualize.py            # all figures
    python3 visualize.py survey     # one or more parts: survey, social-media,
                                    # phd-degree, job-migration, publication,
                                    # patent, commercial-models

Figures are written to ../figures/ as PDF.

Data files (all aggregated, no individual-level records):

    ../data/survey/survey.xlsx                       Figure 1
    ../data/social-media/social-media.xlsx           Figure 2, Figure S7
    ../data/phd-degree/phd-degree.xlsx               Figure 3A, Figures S8-S11
    ../data/job-migration/job-migration.xlsx         Figure 3B-C, Figure S12
    ../data/publication/publication.xlsx             Figure 4A, Figure S13
    ../data/patent/patent.xlsx                       Figures S15-S16 (Figure 4B needs
                                                     the Fang et al. data, see README)
    ../data/commercial-models/commercial-models.xlsx Figure 5

Requires: python >= 3.9, numpy, pandas, matplotlib, openpyxl.
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.gridspec as gridspec
import matplotlib.ticker as mticker
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import openpyxl

# ── Paths ──────────────────────────────────────────────────────────────────────
CODE_DIR = Path(__file__).resolve().parent
DATA = CODE_DIR / '..' / 'data'
FIGURES = CODE_DIR / '..' / 'figures'


def data_file(part):
    """Path of the data workbook of one part: ../data/<part>/<part>.xlsx."""
    return DATA / part / f'{part}.xlsx'


# ── Colour palette (colorbrewer2 Paired-8) ─────────────────────────────────────
COLOR_CN = '#ff7f00'   # orange -> China
COLOR_US = '#1f78b4'   # blue   -> United States
# Neutral gray matched to the mean Rec. 709 luminance of COLOR_CN / COLOR_US,
# so a single-series line has the same grayscale weight as orange and blue.
COLOR_NEUTRAL = '#8d8d8d'
LINE_WIDTH = 5
INT_FMT = mticker.FuncFormatter(lambda v, _: f'{v:,.0f}')

YEARS = list(range(2015, 2026))
CN_LAST_OFFICIAL = 2022   # last year with official Chinese STEM doctoral totals
US_LAST_OFFICIAL = 2024   # last year with official US STEM doctoral totals


# ═══════════════════════════════════════════════════════════════════════════════
#  Reusable plotting helpers
# ═══════════════════════════════════════════════════════════════════════════════

def savefig(fig, filename, hide_top_spine=True, hide_right_spine=True,
            transparent=True, dpi=300, tight_layout=True):
    """Save a figure with the top/right spines hidden on every axis."""
    filename = Path(filename)
    filename.parent.mkdir(parents=True, exist_ok=True)
    if tight_layout:
        fig.tight_layout()
    for ax in fig.axes:
        if hide_top_spine:
            ax.spines['top'].set_visible(False)
        if hide_right_spine:
            ax.spines['right'].set_visible(False)
    fig.savefig(filename, transparent=transparent, dpi=dpi)
    print('wrote', filename.resolve())


def add_subplot_label(ax, label, x=-0.2, y=1.02):
    """Add bold uppercase letter in the upper-left corner (PNAS style)."""
    ax.text(x, y, label, transform=ax.transAxes,
            fontsize=12, fontweight='bold', va='top', ha='left')


def plot_barh(ax, categories, values_cn, values_us, *,
              show_values=True, value_fmt='{:.1f}%', xlim=None,
              xlabel='Percentage', bar_height=0.35):
    """Horizontal bar chart comparing China and the US."""
    y = np.arange(len(categories))
    ax.barh(y - bar_height / 2, values_cn, bar_height,
            color=COLOR_CN, label='China', zorder=3)
    ax.barh(y + bar_height / 2, values_us, bar_height,
            color=COLOR_US, label='United States', zorder=3)

    if show_values:
        x_offset = (xlim[1] - xlim[0]) * 0.005 if xlim else 0.5
        for i, (vc, vu) in enumerate(zip(values_cn, values_us)):
            ax.text(vu + x_offset, i + bar_height / 2, value_fmt.format(vu),
                    va='center', ha='left', fontsize=11, color=COLOR_US)
            ax.text(vc + x_offset, i - bar_height / 2, value_fmt.format(vc),
                    va='center', ha='left', fontsize=11, color=COLOR_CN)

    ax.set_yticks(y)
    ax.set_yticklabels(categories)
    ax.invert_yaxis()
    if xlabel:
        ax.set_xlabel(xlabel)
    if xlim:
        ax.set_xlim(xlim)


def plot_line(ax, x, y_cn, y_us, *,
              xlabel='', ylabel='', title='',
              cn_label='China', us_label='United States',
              marker='', ms=5, legend=True):
    """Line chart comparing China and the US."""
    ax.plot(x, y_cn, color=COLOR_CN, lw=LINE_WIDTH, marker=marker, ms=ms,
            label=cn_label, zorder=3)
    ax.plot(x, y_us, color=COLOR_US, lw=LINE_WIDTH, marker=marker, ms=ms,
            label=us_label, zorder=3)
    if xlabel:
        ax.set_xlabel(xlabel)
    if ylabel:
        ax.set_ylabel(ylabel)
    if title:
        ax.set_title(title)
    if legend:
        ax.legend(loc='best', frameon=False)


def _wrap_label(text, max_chars=25):
    """Insert newlines to wrap long labels."""
    if len(text) <= max_chars:
        return text
    words = text.split()
    lines, current = [], ''
    for w in words:
        if current and len(current) + 1 + len(w) > max_chars:
            lines.append(current)
            current = w
        else:
            current = f'{current} {w}'.strip()
    if current:
        lines.append(current)
    return '\n'.join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure 1 - Survey: AI exposure, AI use and attitudes
# ═══════════════════════════════════════════════════════════════════════════════

def make_figure_survey():
    """5-panel survey figure (A-E) from survey.xlsx."""
    subplot_label_x = -0.5

    wb = openpyxl.load_workbook(data_file('survey'), data_only=True)

    # Sheet F1_A: overall exposure
    rows_a = list(wb['F1_A'].iter_rows(min_row=2, values_only=True))
    heard_us = [r[2] for r in rows_a if r[0] == 'Ever heard of AI' and r[1] == 'U.S.'][0]
    heard_cn = [r[2] for r in rows_a if r[0] == 'Ever heard of AI' and r[1] == 'China'][0]
    used_us = [r[2] for r in rows_a if r[0] == 'Ever used AI' and r[1] == 'U.S.'][0]
    used_cn = [r[2] for r in rows_a if r[0] == 'Ever used AI' and r[1] == 'China'][0]

    # Sheet F1_B+C: exposure by education
    rows_bc = list(wb['F1_B+C'].iter_rows(min_row=2, values_only=True))

    def get_edu(metric, edu, country):
        return [r[3] for r in rows_bc
                if r[0] == metric and r[1] == edu and r[2] == country][0]

    heard_edu_cn = [get_edu('Ever heard of AI', 'College and above', 'China'),
                    get_edu('Ever heard of AI', 'HS or less', 'China')]
    heard_edu_us = [get_edu('Ever heard of AI', 'College and above', 'U.S.'),
                    get_edu('Ever heard of AI', 'HS or less', 'U.S.')]
    used_edu_cn = [get_edu('Ever used AI', 'College and above', 'China'),
                   get_edu('Ever used AI', 'HS or less', 'China')]
    used_edu_us = [get_edu('Ever used AI', 'College and above', 'U.S.'),
                   get_edu('Ever used AI', 'HS or less', 'U.S.')]

    def collect_metric_country(rows):
        labels, cn, us = [], [], []
        seen = set()
        for r in rows:
            metric, country, val = r[0], r[1], r[2]
            if metric not in seen:
                seen.add(metric)
                labels.append(metric)
                cn.append(None)
                us.append(None)
            idx = labels.index(metric)
            if country == 'China':
                cn[idx] = val
            else:
                us[idx] = val
        return labels, cn, us

    att_labels_pct, att_cn_pct, att_us_pct = collect_metric_country(
        list(wb['F1_D'].iter_rows(min_row=2, values_only=True)))
    att_labels_mean, att_cn_mean, att_us_mean = collect_metric_country(
        list(wb['F1_E'].iter_rows(min_row=2, values_only=True)))
    wb.close()

    # Layout: 3 rows x 2 cols. A spans rows 0-1 col 0; B row 0 col 1;
    # C row 1 col 1; D row 2 col 0; E row 2 col 1.
    n_pct = len(att_labels_pct)
    n_mean = len(att_labels_mean)
    fig = plt.figure(figsize=(10, 5))
    gs = gridspec.GridSpec(3, 2,
                           height_ratios=[1, 1, max(n_pct, n_mean)],
                           hspace=0.40, wspace=0.60,
                           left=0.15, right=0.95, top=0.97, bottom=0.08)

    # A: overall exposure
    ax_a = fig.add_subplot(gs[0:2, 0])
    plot_barh(ax_a,
              categories=['Ever heard of AI', 'Ever used AI'],
              values_cn=[heard_cn, used_cn],
              values_us=[heard_us, used_us],
              xlim=(0, 100), xlabel='')
    ax_a.xaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=0))
    ax_a.legend(loc='center right', frameon=True, fancybox=True,
                facecolor=(1, 1, 1, 0.7), edgecolor='lightgrey',
                bbox_to_anchor=(1.0, 0.4))
    add_subplot_label(ax_a, 'A', x=-0.35)

    # B: "Ever heard of AI" by education
    ax_b = fig.add_subplot(gs[0, 1])
    plot_barh(ax_b,
              categories=['Ever heard of AI\n(College and above)',
                          'Ever heard of AI\n(HS or less)'],
              values_cn=heard_edu_cn,
              values_us=heard_edu_us,
              xlim=(0, 100), xlabel='')
    ax_b.xaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=0))
    add_subplot_label(ax_b, 'B', x=subplot_label_x)

    # C: "Ever used AI" by education
    ax_c = fig.add_subplot(gs[1, 1])
    plot_barh(ax_c,
              categories=['Ever used AI\n(College and above)',
                          'Ever used AI\n(HS or less)'],
              values_cn=used_edu_cn,
              values_us=used_edu_us,
              xlim=(0, 100), xlabel='')
    ax_c.xaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=0))
    add_subplot_label(ax_c, 'C', x=subplot_label_x)

    # D: attitudes, percentage items
    ax_d = fig.add_subplot(gs[2, 0])
    plot_barh(ax_d,
              categories=[_wrap_label(l, 18) for l in att_labels_pct],
              values_cn=att_cn_pct,
              values_us=att_us_pct,
              xlim=(0, 100), xlabel='',
              value_fmt='{:.1f}%')
    ax_d.xaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=0))
    ax_d.legend().set_visible(False)
    add_subplot_label(ax_d, 'D', x=-0.35)

    # E: attitudes, mean-value items
    ax_e = fig.add_subplot(gs[2, 1])
    plot_barh(ax_e,
              categories=[_wrap_label(l, 20) for l in att_labels_mean],
              values_cn=att_cn_mean,
              values_us=att_us_mean,
              xlim=(0, 10), xlabel='Level of comfort',
              value_fmt='{:.2f}')
    ax_e.legend().set_visible(False)
    add_subplot_label(ax_e, 'E', x=subplot_label_x)

    savefig(fig, FIGURES / 'figure-1-survey.pdf', tight_layout=False)
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure 2, Figure S7 - Social media attitudes toward AI
# ═══════════════════════════════════════════════════════════════════════════════

def _load_social_media():
    """Daily attitude series (unsmoothed) indexed by date.

    Returns a 3-observation centred moving average along the observation
    order; missing values are skipped inside each window.
    """
    df = pd.read_excel(data_file('social-media'), sheet_name='daily')
    df['date'] = pd.to_datetime(df['date'])
    df = df.set_index('date').sort_index()
    return df.rolling(window=3, center=True, min_periods=1).mean()


def _draw_social_media(series, specs, trend=False, legend_loc='lower right'):
    """Broken-axis social media panel: top [1,2], middle [0,1], bottom [-2,0].

    The series are drawn on the middle panel, which takes most of the height.
    If ``trend`` is True, draw scatter points plus a fitted quadratic curve
    instead of the polyline.
    """
    height_ratios = [1, 20, 1]
    fig = plt.figure(figsize=(5, 4))
    gs = gridspec.GridSpec(3, 1, height_ratios=height_ratios, hspace=0.04,
                           left=0.12, right=0.95, top=0.95, bottom=0.13)
    ax_top = fig.add_subplot(gs[0])
    ax_mid = fig.add_subplot(gs[1], sharex=ax_top)
    ax_bot = fig.add_subplot(gs[2], sharex=ax_top)
    axes = (ax_top, ax_mid, ax_bot)

    for ax in axes:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    def _draw_one(ax, y, spec, label):
        if trend:
            ax.scatter(y.index, y.values, color=spec['color'], s=16, zorder=3)
            if len(y) >= 3:
                xnum = mdates.date2num(pd.to_datetime(y.index))
                coeffs = np.polyfit(xnum, y.values.astype(float), 2)
                xline = np.linspace(xnum.min(), xnum.max(), 50)
                ax.plot(mdates.num2date(xline), np.polyval(coeffs, xline),
                        color=spec['color'], lw=LINE_WIDTH, ls=spec.get('ls', '-'),
                        zorder=4, label=label)
        else:
            ax.plot(y.index, y.values,
                    color=spec['color'], lw=LINE_WIDTH, ls=spec.get('ls', '-'),
                    label=label, zorder=3)

    for spec in specs:
        y = series[spec['column']].dropna()
        _draw_one(ax_mid, y, spec, spec['label'])

    ax_top.set_ylim(1.12, 2.08)
    ax_top.set_yticks([2.0])
    ax_mid.set_ylim(-0.05, 1.05)
    ax_mid.set_yticks([0.0, 0.5, 1.0])
    ax_bot.set_ylim(-3.4, -0.08)
    ax_bot.set_yticks([-2.0])

    yfmt = mticker.FormatStrFormatter('%.1f')
    for ax in axes:
        ax.yaxis.set_major_formatter(yfmt)

    ax_top.spines['bottom'].set_visible(False)
    ax_top.tick_params(axis='x', which='both', bottom=False, labelbottom=False)
    ax_mid.spines['top'].set_visible(False)
    ax_mid.spines['bottom'].set_visible(False)
    ax_mid.tick_params(axis='x', which='both', bottom=False, labelbottom=False)
    ax_bot.spines['top'].set_visible(False)

    # Axis-break slashes, centred on the left spine.
    dx = 0.012
    dy_mid = dx
    dy_top = dx * height_ratios[1] / height_ratios[0]
    dy_bot = dx * height_ratios[1] / height_ratios[2]
    kwargs = dict(color='k', clip_on=False, lw=0.8)
    ax_top.plot((-dx, +dx), (-dy_top, +dy_top), transform=ax_top.transAxes, **kwargs)
    ax_mid.plot((-dx, +dx), (1 - dy_mid, 1 + dy_mid), transform=ax_mid.transAxes, **kwargs)
    ax_mid.plot((-dx, +dx), (-dy_mid, +dy_mid), transform=ax_mid.transAxes, **kwargs)
    ax_bot.plot((-dx, +dx), (1 - dy_bot, 1 + dy_bot), transform=ax_bot.transAxes, **kwargs)

    ax_bot.set_xlim(series.index.min() - pd.Timedelta(days=18), None)
    ax_bot.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[3, 7, 11]))
    _months = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
    ax_bot.xaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _pos: f'{_months[mdates.num2date(x).month - 1]} {mdates.num2date(x).year}'))

    fig.text(0.02, 0.5, 'Social media attitude toward AI',
             va='center', ha='center', rotation='vertical', fontsize=10)
    ax_top.text(0.01, 2.0, 'Excited', fontsize=9, color='grey', va='center',
                transform=ax_top.get_yaxis_transform())
    ax_mid.text(0.01, 0.0, 'Neutral', fontsize=9, color='grey', va='center',
                transform=ax_mid.get_yaxis_transform())
    ax_bot.text(0.01, -2.0, 'Concerned', fontsize=9, color='grey', va='center',
                transform=ax_bot.get_yaxis_transform())

    handles, labels = ax_mid.get_legend_handles_labels()
    # loc is the corner of the legend box that attaches to bbox_to_anchor.
    legend_anchor = {'lower right': (0.95, 0.10), 'upper left': (0.15, 0.95)}[legend_loc]
    fig.legend(handles, labels, loc=legend_loc, frameon=False,
               bbox_to_anchor=legend_anchor, bbox_transform=fig.transFigure)
    return fig


def make_figure_social_media():
    """Figure 2 (main text) and Figure S7 from social-media.xlsx."""
    series = _load_social_media()

    # Figure 2: user-level average attitude, scatter + quadratic trend.
    fig = _draw_social_media(series, [
        dict(column='weibo-deepseek', color=COLOR_CN, label='China (Weibo)'),
        dict(column='twitter-deepseek', color=COLOR_US, label='United States (Twitter)'),
    ], trend=True)
    savefig(fig, FIGURES / 'figure-2-social-media.pdf', tight_layout=False)
    plt.close(fig)

    # Figure S7: robustness to the LLM classifier (Twitter posts classified
    # by DeepSeek and, independently, by GPT-5-mini).
    fig = _draw_social_media(series, [
        dict(column='weibo-deepseek', color=COLOR_CN, ls='-',
             label='China (Weibo, DeepSeek)'),
        dict(column='twitter-deepseek', color=COLOR_US, ls='-',
             label='United States (Twitter, DeepSeek)'),
        dict(column='twitter-gpt', color=COLOR_US, ls=':',
             label='United States (Twitter, GPT)'),
    ], legend_loc='upper left')
    savefig(fig, FIGURES / 'figure-S7-social-media-llm-robustness.pdf', tight_layout=False)
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure 3 - Human capital: PhD dissertations (A) and AI job transitions (B, C)
# ═══════════════════════════════════════════════════════════════════════════════

CALIBER_BASELINE = 'Abstract, with acronyms'


def _load_degree_estimates(caliber=CALIBER_BASELINE):
    """Estimated AI-related doctorates for one query caliber.

    Returns (cn, us): DataFrames indexed by year with columns
    point, upper, lower, status.
    """
    df = pd.read_excel(data_file('phd-degree'), sheet_name='estimates')
    df = df[df['caliber'] == caliber]
    if df.empty:
        raise ValueError(f'caliber not found in phd-degree.xlsx: {caliber}')
    cn = df[df['country'] == 'China'].set_index('year').sort_index()
    us = df[df['country'] == 'United States'].set_index('year').sort_index()
    return cn, us


def _load_degree_shares():
    """In-database AI share of STEM dissertations (fractions), by country."""
    df = pd.read_excel(data_file('phd-degree'), sheet_name='in_database_shares')
    cn = df[df['country'] == 'China'].set_index('year').sort_index()
    us = df[df['country'] == 'United States'].set_index('year').sort_index()
    return cn, us


def _load_jobs_odds():
    """Yearly odds ratios of US-to-China vs China-to-US AI job transitions."""
    df = pd.read_excel(data_file('job-migration'), sheet_name='odds_ratio')
    return df[['year', 'all_positions']], df[['year', 'managerial_positions']]


def _plot_split_line(ax, years, pt, last_official, color, label):
    """Solid through last official year; dotted for estimates."""
    years = np.asarray(years)
    pt = np.asarray(pt)
    i_off = int(np.where(years == last_official)[0][0])
    i_dot = min(i_off, len(years) - 2)
    ax.plot(years[: i_dot + 1], pt[: i_dot + 1],
            color=color, lw=LINE_WIDTH, solid_capstyle='round', zorder=3,
            label=label)
    ax.plot(years[i_dot:], pt[i_dot:],
            color=color, lw=LINE_WIDTH, zorder=4,
            linestyle=(0, (1.5, 0.5)), dash_capstyle='butt')


def _draw_degree_line(ax, cn, us, ylabel, *, legend=False, label=None):
    """PhD counts: official solid, estimated dotted."""
    _plot_split_line(ax, cn.index, cn['point'], CN_LAST_OFFICIAL, COLOR_CN, 'China')
    _plot_split_line(ax, us.index, us['point'], US_LAST_OFFICIAL, COLOR_US, 'United States')
    ax.set_ylabel(ylabel)
    ax.yaxis.set_major_formatter(INT_FMT)
    ax.set_ylim(bottom=0)
    if legend:
        ax.legend(loc='best', frameon=False)
    if label:
        add_subplot_label(ax, label, x=-0.3)


def _draw_jobs_odds(ax, df, column, ylabel, *, ymax, label=None):
    """Single-series odds-ratio time series."""
    ax.plot(df['year'], df[column], color=COLOR_NEUTRAL, lw=LINE_WIDTH, zorder=3)
    ax.set_ylabel(ylabel)
    ax.set_xticks(range(2015, 2026, 2))
    ax.set_xlim(2014.5, 2025.5)
    ax.set_ylim(0, ymax * 1.08)
    if label:
        add_subplot_label(ax, label, x=-0.3)


def make_figure_human_capital():
    """Figure 3: PhD production (A) + job-transition odds ratios (B, C)."""
    deg_cn, deg_us = _load_degree_estimates()
    df_all, df_mgr = _load_jobs_odds()
    ymax = max(df_all['all_positions'].max(), df_mgr['managerial_positions'].max())

    fig, axes = plt.subplots(3, 1, figsize=(5, 8))
    _draw_degree_line(
        axes[0], deg_cn, deg_us,
        'Number of AI-related PhD\ndissertations in STEM',
        legend=True, label='A')
    _draw_jobs_odds(
        axes[1], df_all, 'all_positions',
        'Odds ratio of AI job transitions\n(all jobs)',
        ymax=ymax, label='B')
    _draw_jobs_odds(
        axes[2], df_mgr, 'managerial_positions',
        'Odds ratio of AI job transitions\n(AI managerial jobs)',
        ymax=ymax, label='C')

    savefig(fig, FIGURES / 'figure-3-human-capital.pdf')
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure 4 - Knowledge capital: publications (A) and patents (B)
# ═══════════════════════════════════════════════════════════════════════════════

def _load_projected_publications():
    """Projected annual AI publications in CS (Stanford AI Index, adjusted
    for publications with unknown country affiliation)."""
    df = pd.read_excel(data_file('publication'), sheet_name='stanford_projected')
    df = df.set_index('year').sort_index()
    return df.index.astype(int), df['China'], df['United States']


def _load_patent_grants():
    """AI patent grants by granting office and grant year, 2015-2023,
    aggregated from Fang et al. (NBER Working Paper 35022).

    That data set is not redistributed here. To draw Figure 4B, add a sheet
    `fang_grants` with columns year, China, United States to patent.xlsx
    (see README). Returns None when the sheet is absent.
    """
    xlsx = data_file('patent')
    if 'fang_grants' not in pd.ExcelFile(xlsx).sheet_names:
        print('patent.xlsx has no sheet fang_grants: Figure 4 is drawn without panel B')
        return None
    df = pd.read_excel(xlsx, sheet_name='fang_grants')
    df = df[df['year'] >= 2015].sort_values('year')
    return df['year'].astype(int).tolist(), df['China'].tolist(), df['United States'].tolist()


def _draw_count_line(ax, x, y_cn, y_us, ylabel, *, legend=False, label=None,
                     ylim_bottom=0):
    """China/US count time series."""
    plot_line(ax, x=x, y_cn=y_cn, y_us=y_us, ylabel=ylabel, legend=legend)
    ax.yaxis.set_major_formatter(INT_FMT)
    if ylim_bottom is not None:
        ax.set_ylim(bottom=ylim_bottom)
    if label:
        add_subplot_label(ax, label, x=-0.3)


def make_figure_knowledge_capital():
    """Figure 4: projected publications (A) + patent grants (B)."""
    papers_years, papers_cn, papers_us = _load_projected_publications()
    patents = _load_patent_grants()

    fig, axes = plt.subplots(2, 1, figsize=(5, 5.5))
    _draw_count_line(
        axes[0], papers_years, papers_cn.values, papers_us.values,
        'Projected annual number of\nscientific publications',
        legend=True, label='A')
    if patents is None:
        axes[1].set_axis_off()
        axes[1].text(0.5, 0.5, 'Panel B: AI patent grants\n(data from Fang et al., not redistributed)',
                     ha='center', va='center', fontsize=9, color='grey',
                     transform=axes[1].transAxes)
    else:
        pat_years, pat_cn, pat_us = patents
        _draw_count_line(
            axes[1], pat_years, pat_cn, pat_us,
            'Number of AI-related patents',
            label='B', ylim_bottom=None)

    savefig(fig, FIGURES / 'figure-4-knowledge-capital.pdf')
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure 5 - AI models: performance over time
# ═══════════════════════════════════════════════════════════════════════════════

def _load_models():
    """Model performance vs release date (Artificial Analysis Intelligence Index)."""
    df = pd.read_excel(data_file('commercial-models'), sheet_name='Data')
    df = df.dropna(subset=['intelligence_index', 'release_date']).copy()
    df['release_date'] = pd.to_datetime(df['release_date'])
    return df


def make_figure_models():
    """Scatter: model performance vs release date.

    Colour encodes country (orange = China, blue = US). Marker fill encodes
    licence: filled = proprietary, hollow = open-weight.
    """
    df = _load_models()
    fig, ax = plt.subplots(figsize=(5.5, 3.8))

    styles = [
        ('China', 'Open', 'none', COLOR_CN),
        ('United States', 'Open', 'none', COLOR_US),
        ('China', 'Closed', COLOR_CN, COLOR_CN),
        ('United States', 'Closed', COLOR_US, COLOR_US),
    ]
    for country, licence, facecolor, edgecolor in styles:
        sub = df[(df['country'] == country) & (df['open_closed'] == licence)]
        ax.scatter(sub['release_date'], sub['intelligence_index'],
                   facecolors=facecolor, edgecolors=edgecolor,
                   s=42, linewidths=1.1, zorder=3)

    ax.set_xlim(pd.Timestamp('2023-01-01'), pd.Timestamp('2026-11-01'))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    ax.set_ylabel('Model performance')
    ax.set_ylim(0, 72)
    ax.yaxis.set_major_locator(mticker.MultipleLocator(10))
    ax.legend(
        handles=[
            Line2D([0], [0], marker='o', color='none',
                   markerfacecolor=COLOR_CN, markeredgecolor=COLOR_CN,
                   markersize=7, label='China, proprietary'),
            Line2D([0], [0], marker='o', color='none',
                   markerfacecolor='none', markeredgecolor=COLOR_CN,
                   markersize=7, label='China, open-weight'),
            Line2D([0], [0], marker='o', color='none',
                   markerfacecolor=COLOR_US, markeredgecolor=COLOR_US,
                   markersize=7, label='United States, proprietary'),
            Line2D([0], [0], marker='o', color='none',
                   markerfacecolor='none', markeredgecolor=COLOR_US,
                   markersize=7, label='United States, open-weight'),
        ],
        loc='upper left', frameon=False, fontsize=8)

    # Labels sit upper-left of their point (ha=right, va=bottom) except
    # DeepSeek V4, which sits upper-right. Offsets are in points.
    annotations = [
        ('Qwen Chat 72B', 'Qwen Chat 72B', (-4, 32), 'right'),
        ('Claude 3 Opus', 'Claude 3 Opus', (-5, 18), 'right'),
        ("DeepSeek R1 (Jan '25)", 'DeepSeek R1', (-8, 4), 'right'),
        ('Claude 3.7 Sonnet (Reasoning)', 'Claude 3.7', (-6, 9), 'right'),
        ('Gemini 3.1 Pro Preview', 'Gemini 3.1', (-22, 8), 'right'),
        ('GPT-5.5 (xhigh)', 'GPT-5.5', (-6, 8), 'right'),
        ('DeepSeek V4 Pro (max)', 'DeepSeek V4', (6, 8), 'left'),
        ('Claude Opus 5 (max)', 'Claude Opus 5', (-10, 26), 'right'),
        ('GLM-5.3 (max)', 'GLM-5.3', (-4, 22), 'right'),
    ]
    for name, text, xytext, ha in annotations:
        row = df.loc[df['model_name'] == name]
        if row.empty:
            raise ValueError(f'Model not found for annotation: {name}')
        row = row.iloc[0]
        colour = COLOR_CN if row['country'] == 'China' else COLOR_US
        ax.annotate(
            text,
            xy=(row['release_date'], row['intelligence_index']),
            xytext=xytext, textcoords='offset points',
            fontsize=6.5, ha=ha, va='bottom', color=colour, zorder=4,
            arrowprops=dict(arrowstyle='-', color=colour, lw=0.6,
                            shrinkA=0, shrinkB=4))

    savefig(fig, FIGURES / 'figure-5-models.pdf')
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figures S8-S11 - AI-related doctoral dissertations (SI Appendix)
# ═══════════════════════════════════════════════════════════════════════════════

SI_DEGREE_US = '#1f77b4'
SI_DEGREE_CN = '#ff7f0e'
SI_DEGREE_GRAY = '#bdbdbd'


def _si_degree_style(legend_size=9):
    plt.rcParams.update({
        'font.family': 'sans-serif',
        'font.size': 11,
        'axes.labelsize': 11,
        'axes.titlesize': 12,
        'legend.fontsize': legend_size,
        'axes.linewidth': 0.8,
        'pdf.fonttype': 42,
        'ps.fonttype': 42,
    })


def _si_degree_axes(ax):
    ax.set_xticks(YEARS)
    ax.set_xlim(2014.6, 2025.4)
    ax.grid(True, color='#d0d0d0', linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_color('#444444')


def make_figure_degree_share():
    """Figure S8: in-database AI share of STEM dissertations
    (abstract field, vocabulary with acronyms)."""
    with plt.rc_context():
        _si_degree_style(legend_size=10)
        cn, us = _load_degree_shares()
        fig, ax = plt.subplots(figsize=(7.4, 4.6))
        ax.plot(YEARS, 100 * us.loc[YEARS, CALIBER_BASELINE], color=SI_DEGREE_US,
                marker='o', markersize=6, linewidth=2, label='United States (ProQuest)', zorder=3)
        ax.plot(YEARS, 100 * cn.loc[YEARS, CALIBER_BASELINE], color=SI_DEGREE_CN,
                marker='o', markersize=6, linewidth=2, label='China (CNKI)', zorder=3)
        ax.set_title('In-database AI share of STEM doctoral dissertations\n(Abstract + with acronyms)')
        ax.set_xlabel('Year')
        ax.set_ylabel('AI share of STEM dissertations (%)')
        ax.set_ylim(0, 26)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.0f}'))
        _si_degree_axes(ax)
        ax.legend(loc='upper left', frameon=True, fancybox=False, edgecolor='#cccccc')
        fig.tight_layout()
        FIGURES.mkdir(parents=True, exist_ok=True)
        out = FIGURES / 'figure-S8-phd-degree-in-database-share.pdf'
        fig.savefig(out, bbox_inches='tight')
        print('wrote', out.resolve())
        plt.close(fig)


def _si_degree_band(ax, years, lo, hi, i0, edge):
    xs = years[i0:]
    ax.fill_between(xs, lo[i0:], hi[i0:], color=SI_DEGREE_GRAY, alpha=0.45,
                    linewidth=0, zorder=1)
    ax.plot(xs, lo[i0:], linestyle='--', color=edge, linewidth=1.4, zorder=2)
    ax.plot(xs, hi[i0:], linestyle='--', color=edge, linewidth=1.4, zorder=2)


def _si_degree_split(ax, years, y, i_last_official, color):
    ax.plot(years[: i_last_official + 1], y[: i_last_official + 1],
            color=color, marker='o', markersize=6, linewidth=2, zorder=3, linestyle='-')
    ax.plot(years[i_last_official:], y[i_last_official:],
            color=color, linewidth=2, zorder=3, linestyle='--')
    if i_last_official + 1 < len(years):
        ax.plot(years[i_last_official + 1:], y[i_last_official + 1:],
                color=color, marker='o', markersize=6, linestyle='None', zorder=4,
                markerfacecolor='white', markeredgecolor=color, markeredgewidth=1.4)


def make_figure_degree_estimates():
    """Figures S9-S11: estimated AI-related doctorates, three alternative
    query calibers (the baseline caliber is Figure 3A)."""
    calibers = [
        ('S9', 'Abstract, without acronyms', 'abstract-without-acronyms'),
        ('S10', 'Keywords, with acronyms', 'keywords-with-acronyms'),
        ('S11', 'Keywords, without acronyms', 'keywords-without-acronyms'),
    ]
    i_cn = YEARS.index(CN_LAST_OFFICIAL)
    i_us = YEARS.index(US_LAST_OFFICIAL)
    for number, caliber, slug in calibers:
        with plt.rc_context():
            _si_degree_style(legend_size=9)
            cn, us = _load_degree_estimates(caliber)
            cn_lo, cn_pt, cn_hi = (cn.loc[YEARS, c].tolist() for c in ('lower', 'point', 'upper'))
            us_lo, us_pt, us_hi = (us.loc[YEARS, c].tolist() for c in ('lower', 'point', 'upper'))

            fig, ax = plt.subplots(figsize=(7.4, 4.6))
            _si_degree_band(ax, YEARS, cn_lo, cn_hi, i_cn, SI_DEGREE_CN)
            _si_degree_band(ax, YEARS, us_lo, us_hi, i_us, SI_DEGREE_US)
            _si_degree_split(ax, YEARS, us_pt, i_us, SI_DEGREE_US)
            _si_degree_split(ax, YEARS, cn_pt, i_cn, SI_DEGREE_CN)
            # 2025 United States: explicit error bar because only that year
            # is estimated for the US.
            ax.errorbar([2025], [us_pt[-1]],
                        yerr=[[us_pt[-1] - us_lo[-1]], [us_hi[-1] - us_pt[-1]]],
                        fmt='none', ecolor=SI_DEGREE_US, elinewidth=1.4,
                        capsize=5, capthick=1.4, zorder=5)
            ax.set_title(f'Estimated AI-related doctorates\n({caliber})')
            ax.set_xlabel('Year')
            ax.set_ylabel('Estimated AI-related doctorates')
            ax.set_ylim(0, None)
            _si_degree_axes(ax)
            ax.legend(handles=[
                Line2D([0], [0], color=SI_DEGREE_US, marker='o', linewidth=2,
                       label='United States (official)'),
                Line2D([0], [0], color=SI_DEGREE_US, marker='o', linewidth=2, linestyle='--',
                       markerfacecolor='white', markeredgecolor=SI_DEGREE_US,
                       label='United States (estimated)'),
                Line2D([0], [0], color=SI_DEGREE_CN, marker='o', linewidth=2,
                       label='China (official)'),
                Line2D([0], [0], color=SI_DEGREE_CN, marker='o', linewidth=2, linestyle='--',
                       markerfacecolor='white', markeredgecolor=SI_DEGREE_CN,
                       label='China (estimated)'),
                Patch(facecolor=SI_DEGREE_GRAY, alpha=0.45, edgecolor='none',
                      label='Estimation interval'),
            ], loc='upper left', frameon=True, fancybox=False, edgecolor='#cccccc')
            fig.tight_layout()
            FIGURES.mkdir(parents=True, exist_ok=True)
            out = FIGURES / f'figure-{number}-phd-degree-{slug}.pdf'
            fig.savefig(out, bbox_inches='tight')
            print('wrote', out.resolve())
            plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure S12 - AI-related career moves between the US and China (SI Appendix)
# ═══════════════════════════════════════════════════════════════════════════════

def make_figure_job_moves():
    """Figure S12: annual counts of US-China AI career moves, all positions (A)
    and managerial positions (B). US-to-China counts are reweighted for the
    change in LinkedIn activity in China after 2021."""
    xlsx = data_file('job-migration')
    df_all = pd.read_excel(xlsx, sheet_name='moves_all')
    df_mgr = pd.read_excel(xlsx, sheet_name='moves_managerial')

    fig, axes = plt.subplots(2, 1, figsize=(5, 5.5))
    for ax, df, ylabel, letter, legend in (
            (axes[0], df_all, 'Number of AI professional moves\n(all positions)', 'A', True),
            (axes[1], df_mgr, 'Number of AI professional moves\n(managerial positions)', 'B', False)):
        plot_line(ax,
                  x=df['year'].values,
                  y_us=df['China_to_US'].values,
                  y_cn=df['US_to_China_weighted'].values,
                  ylabel=ylabel,
                  us_label='China to US',
                  cn_label='US to China',
                  legend=legend)
        ax.yaxis.set_major_formatter(INT_FMT)
        ax.set_xticks(range(2015, 2026, 2))
        ax.set_xlim(2014.5, 2025.5)
        ax.set_ylim(bottom=0)
        add_subplot_label(ax, letter, x=-0.3)

    savefig(fig, FIGURES / 'figure-S12-job-migration-moves.pdf')
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figure S13 - OpenAlex robustness check (SI Appendix)
# ═══════════════════════════════════════════════════════════════════════════════

def make_figure_openalex():
    """Figure S13: yearly counts of AI papers from China and the US in OpenAlex,
    2015-2024 (title keyword matching, corresponding/first-author country)."""
    counts = pd.read_excel(data_file('publication'), sheet_name='openalex')
    counts = counts.set_index('year').sort_index().loc[2015:2024]

    # Colours of the original SI figure (colorbrewer2 Oranges / Blues).
    color_cn, color_us = '#fd8d3c', '#6baed6'
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(counts.index, counts['CN'], lw=5, label='China', marker='o', color=color_cn)
    ax.plot(counts.index, counts['US'], lw=5, label='United States', marker='o', color=color_us)
    ax.set(title='Yearly Counts of AI Papers from China and the US',
           xlabel='Year', ylabel='Yearly number of Papers')
    ax.yaxis.set_major_formatter(INT_FMT)
    ax.legend(loc=0, frameon=False, fancybox=True)
    savefig(fig, FIGURES / 'figure-S13-publication-openalex.pdf', transparent=False)
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Figures S15-S16 - Lens.org patent families (SI Appendix)
# ═══════════════════════════════════════════════════════════════════════════════

def make_figure_patent_lens():
    """Figures S15-S16: AI-related simple patent families by priority
    jurisdiction, 2015-2024: all families (S15) and families cited by
    another patent (S16)."""
    df = pd.read_excel(data_file('patent'), sheet_name='lens_families')
    df = df.set_index('year').sort_index().loc[2015:2024]
    color_cn, color_us = '#F87800', '#1878B0'
    specs = [
        ('S15', 'China_all', 'United_States_all', 'all'),
        ('S16', 'China_cited', 'United_States_cited', 'cited'),
    ]
    with plt.rc_context({'font.family': 'sans-serif',
                         'font.sans-serif': ['Arial', 'Helvetica', 'DejaVu Sans'],
                         'axes.unicode_minus': False}):
        for number, col_cn, col_us, slug in specs:
            x = df.index.to_numpy()
            fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=200)
            fig.patch.set_facecolor('white')
            ax.plot(x, df[col_cn].to_numpy(), color=color_cn, lw=1.8, marker='o', ms=5.5, zorder=3)
            ax.plot(x, df[col_us].to_numpy(), color=color_us, lw=1.8, marker='o', ms=5.5, zorder=3)
            ax.set_xlim(2015.2, 2024.6)
            ax.set_xticks([2016, 2018, 2020, 2022, 2024])
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _p: f'{int(v):,}'))
            ax.set_ylabel('Number of AI-related\npatent families', fontsize=10)
            ax.set_ylim(bottom=0)
            ax.set_facecolor('white')
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_color('black')
                spine.set_linewidth(0.9)
            ax.tick_params(axis='both', direction='out', length=4, width=0.8, colors='black')
            ax.grid(False)
            ax.text(-0.12, 1.04, number, transform=ax.transAxes, fontsize=12,
                    fontweight='bold', va='bottom', ha='left')
            ax.legend(handles=[
                Line2D([0], [0], color=color_cn, marker='o', lw=1.8, ms=5.5, label='China'),
                Line2D([0], [0], color=color_us, marker='o', lw=1.8, ms=5.5, label='United States'),
            ], loc='upper left', frameon=False, fontsize=9, handlelength=1.4, borderaxespad=0.4)
            fig.tight_layout()
            FIGURES.mkdir(parents=True, exist_ok=True)
            out = FIGURES / f'figure-{number}-patent-lens-{slug}.pdf'
            fig.savefig(out, dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.08)
            print('wrote', out.resolve())
            plt.close(fig)


# ═══════════════════════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════════════════════

PARTS = {
    'survey': [make_figure_survey],
    'social-media': [make_figure_social_media],
    'phd-degree': [make_figure_degree_share, make_figure_degree_estimates],
    'job-migration': [make_figure_job_moves],
    'publication': [make_figure_openalex],
    'patent': [make_figure_patent_lens],
    'commercial-models': [make_figure_models],
}
# Figures that combine several parts.
COMBINED = {
    'figure-3-human-capital': (make_figure_human_capital, {'phd-degree', 'job-migration'}),
    'figure-4-knowledge-capital': (make_figure_knowledge_capital, {'publication', 'patent'}),
}

if __name__ == '__main__':
    selected = sys.argv[1:] or list(PARTS)
    unknown = [p for p in selected if p not in PARTS]
    if unknown:
        sys.exit(f'unknown part(s): {unknown}; choose from {list(PARTS)}')
    FIGURES.mkdir(parents=True, exist_ok=True)
    for part in selected:
        for func in PARTS[part]:
            func()
    for func, parts in COMBINED.values():
        if parts & set(selected):
            func()
    print(f'Done. Figures saved to {FIGURES.resolve()}')
