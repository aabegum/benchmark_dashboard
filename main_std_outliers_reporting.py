
#! /usr/bin/env python
"""
This program processes data from various distribution companies and generates a
PowerPoint presentation with graphs comparing their performance, including outlier
detection using STD, IQR, and MAD methods.

Author: Onur Arıkan
Date: 2024-08-20
"""
"""
SCRIPT SUMMARY: CONSENSUS OUTLIER DETECTION
-------------------------------------------
Filename: std_report_outlier_main2_consensus.py

This script processes distribution company data to generate a benchmark report 
with high-confidence outlier detection.

Key Features:
- Consensus Logic: Implements a 'find_consensus_outliers' function that only 
  flags data points as outliers if they are detected by BOTH:
    1. IQR (Interquartile Range)
    2. MAD (Median Absolute Deviation)
  This reduces false positives compared to using a single method.

- Reporting: The generated PowerPoint text box ("Bulgu") highlights the 
  Intersection (Consensus) count, providing a clearer signal of genuine 
  performance deviations.

- Visualization: Generates standard scatter, stacked bar, and overlayed graphs.
"""
# Standard library imports
import logging
from pathlib import Path
import re
# Third-party library imports
from matplotlib.colors import ListedColormap
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import seaborn as sns
import yaml

# Local imports
from utils import wrap_text_over_words

# Configure seaborn
sns.set_style(style="whitegrid")
sns.set_palette("muted")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Define regular expression patterns
ILLEGAL_WINDOWS_PATH_CHARACTERS = re.compile(r'[\\/:*?"<>|\n]')
APG_NO_PATTERN = re.compile(r'(\w+\.\d+)')
ENCODING = 'utf-8'

# Define directory paths
CODING_DIRECTORY = Path(__file__).parent
MAIN_DIRECTORY = CODING_DIRECTORY.parent
GRAPHICS_DIRECTORY = MAIN_DIRECTORY / "Grafikler"
REPORTS_DIRECTORY = MAIN_DIRECTORY / "Raporlar"

# Define file paths
CONFIG_PATH = CODING_DIRECTORY / "config.yaml"
PRESENTATION_INTRO_TEMPLATE_PATH = CODING_DIRECTORY / "presentation_intro_template.txt"

# Load YAML configuration
with CONFIG_PATH.open(encoding=ENCODING) as config_file:
    config = yaml.safe_load(config_file)

MASTER_FILE = MAIN_DIRECTORY / config['MASTER_FILE']
TEMPLATE_PATH = MAIN_DIRECTORY / config['TEMPLATE_PATH']

COMPANY_GROUPS = config['COMPANY_GROUPS']
COMPANY_GROUPS_EXCLUDED_FROM_REPORT = config['COMPANY_GROUPS_EXCLUDED_FROM_REPORT']

# Validate company groups
if set(COMPANY_GROUPS_EXCLUDED_FROM_REPORT) - set(COMPANY_GROUPS):
    raise SystemExit("COMPANY_GROUPS_EXCLUDED_FROM_REPORT contains companies that are not in COMPANY_GROUPS. Please check the config.yaml file.")

NUM_OF_COMPANIES = sum(len(companies) for companies in COMPANY_GROUPS.values())
COMPANIES_RANGE = np.arange(1, NUM_OF_COMPANIES + 1)

REPORT_TYPE_CHOICES = "yariyillik", "yillik","cokyillik"
REPORT_YEAR = config['REPORT_YEAR']
REPORT_TYPE = config['REPORT_TYPE']
print(f"Generating {REPORT_TYPE} report for year {REPORT_YEAR}...")
SIGMA: int = config['SIGMA']
IQR_FACTOR: float = config.get('IQR_FACTOR', 1.5)
MAD_THRESHOLD: float = config.get('MAD_THRESHOLD', 3.5)
START_COL = 3
END_COL = START_COL + NUM_OF_COMPANIES
DEFAULT_DECIMAL_DIGITS = config['DEFAULT_DECIMAL_DIGITS']
WORD_WRAP_LIMIT = config['WORD_WRAP_LIMIT']

# Define company color indicators
GROUP_COMPANY_INDICATOR = 0
RIVAL_COMPANY_INDICATOR = 1

# Plot parameters
ANNOTATION_OFFSET_PIXELS = tuple(config['ANNOTATION_OFFSET_PIXELS'])
ANNOTATION_FONT_SIZE = config['ANNOTATION_FONT_SIZE']
HORIZONTAL_MEAN_COLOR = config['HORIZONTAL_MEAN_COLOR']
HORIZONTAL_MEAN_ALPHA = config['HORIZONTAL_MEAN_ALPHA']
FONT_SIZE = config['FONT_SIZE']
OVERLAY_GRAPH_COLOR_MAP = ListedColormap(config['OVERLAY_GRAPH_COLOR_MAP'])
OVERLAY_GRAPH_BAR_COLOR = config['OVERLAY_GRAPH_BAR_COLOR']

def filtered_mean_with_outliers(row: pd.Series, start_col: int, end_col: int, sigma: float = 2) -> dict:
    """
    Calculate the mean of a row within a specified number of standard deviations
    and identify outliers in the specified columns.

    Parameters:
    row (pd.Series): A row from a pandas DataFrame.
    start_col (int): Starting column index for company data.
    end_col (int): Ending column index (exclusive).
    sigma (float, optional): Number of standard deviations. Default is 2.

    Returns:
    dict: Contains filtered_mean, outliers, and outlier_count.
    """
    data = row[start_col:end_col]
    valid_data = data.dropna()
    if len(valid_data) < 2:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Insufficient valid data ({len(valid_data)} non-NaN values)")
        return {'filtered_mean': np.nan, 'outliers': [], 'outlier_count': 0}
    mean = valid_data.mean()
    std = valid_data.std()
    if std == 0 or np.isnan(std):
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Zero or NaN standard deviation")
        return {'filtered_mean': mean if not np.isnan(mean) else np.nan, 'outliers': [], 'outlier_count': 0}
    lower_bound = mean - sigma * std
    upper_bound = mean + sigma * std
    outlier_mask = (data < lower_bound) | (data > upper_bound)
    outliers = [(col, data[col]) for col in data.index[outlier_mask] if outlier_mask[col] and not np.isnan(data[col])]
    filtered_data = data[(data >= lower_bound) & (data <= upper_bound) & data.notna()]
    if filtered_data.empty:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: All values filtered out (STD method)")
        return {'filtered_mean': mean, 'outliers': outliers, 'outlier_count': len(outliers)}
    filtered_mean = filtered_data.mean()
    return {'filtered_mean': filtered_mean, 'outliers': outliers, 'outlier_count': len(outliers)}

def filtered_mean_with_outliers_iqr(row: pd.Series, start_col: int, end_col: int, iqr_factor: float = 1.5) -> dict:
    """
    Calculate the mean of a row using IQR-based outlier detection and identify outliers.

    Parameters:
    row (pd.Series): A row from a pandas DataFrame.
    start_col (int): Starting column index.
    end_col (int): Ending column index (exclusive).
    iqr_factor (float, opstional): Multiplier for IQR bounds. Default is 1.5.

    Returns:
    dict: Contains filtered_mean, outliers, and outlier_count.
    """
    data = row[start_col:end_col]
    valid_data = data.dropna()
    if len(valid_data) < 4:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Insufficient valid data for IQR ({len(valid_data)} non-NaN values)")
        return {'filtered_mean': np.nan, 'outliers': [], 'outlier_count': 0}
    q1 = valid_data.quantile(0.25)
    q3 = valid_data.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0 or np.isnan(iqr):
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Zero or NaN IQR")
        return {'filtered_mean': valid_data.mean() if not valid_data.empty else np.nan, 'outliers': [], 'outlier_count': 0}
    lower_bound = q1 - iqr_factor * iqr
    upper_bound = q3 + iqr_factor * iqr
    outlier_mask = (data < lower_bound) | (data > upper_bound)
    outliers = [(col, data[col]) for col in data.index[outlier_mask] if outlier_mask[col] and not np.isnan(data[col])]
    filtered_data = data[(data >= lower_bound) & (data <= upper_bound) & data.notna()]
    if filtered_data.empty:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: All values filtered out (IQR method)")
        return {'filtered_mean': valid_data.mean(), 'outliers': outliers, 'outlier_count': len(outliers)}
    filtered_mean = filtered_data.mean()
    return {'filtered_mean': filtered_mean, 'outliers': outliers, 'outlier_count': len(outliers)}

def filtered_mean_with_outliers_mad(row: pd.Series, start_col: int, end_col: int, mad_threshold: float = 3.5) -> dict:
    """
    Calculate the mean using MAD-based outlier detection.

    Parameters:
    row (pd.Series): A row from a pandas DataFrame.
    start_col (int): Starting column index.
    end_col (int): Ending column index (exclusive).
    mad_threshold (float, optional): Threshold for modified Z-score. Default is 3.5.

    Returns:
    dict: Contains filtered_mean, outliers, and outlier_count.
    """
    data = row[start_col:end_col]
    valid_data = data.dropna()
    if len(valid_data) < 2:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Insufficient valid data for MAD ({len(valid_data)} non-NaN values)")
        return {'filtered_mean': np.nan, 'outliers': [], 'outlier_count': 0}
    median = valid_data.median()
    mad = np.median(np.abs(valid_data - median))
    if mad == 0 or np.isnan(mad):
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: Zero or NaN MAD")
        return {'filtered_mean': valid_data.mean() if not valid_data.empty else np.nan, 'outliers': [], 'outlier_count': 0}
    modified_z = 0.6745 * (data - median) / mad
    outlier_mask = np.abs(modified_z) > mad_threshold
    outliers = [(col, data[col]) for col in data.index[outlier_mask] if outlier_mask[col] and not np.isnan(data[col])]
    filtered_data = data[~outlier_mask & data.notna()]
    if filtered_data.empty:
        logger.warning(f"APG No {row.get('APG No', 'Unknown')}: All values filtered out (MAD method)")
        return {'filtered_mean': valid_data.mean(), 'outliers': outliers, 'outlier_count': len(outliers)}
    filtered_mean = filtered_data.mean()
    return {'filtered_mean': filtered_mean, 'outliers': outliers, 'outlier_count': len(outliers)}

def generate_presentation_intro_text(company_list: list[str]) -> str:
    """
    Generate the introductory text for the presentation based on the company group.
    """
    presentation_text_template = PRESENTATION_INTRO_TEMPLATE_PATH.read_text(encoding=ENCODING)
    company_lines = []
    for i, company in enumerate(company_list, start=1):
        line = f"- {i} numaralı Şirket, {company}'ı"
        if i == len(company_list):
            line += " temsil etmekte iken"
        company_lines.append(line)
    company_lines.append(f"{num_of_group_companies + 1} - {NUM_OF_COMPANIES}")
    company_enumeration_text = "\n".join(company_lines)
    return presentation_text_template.format(
        company_text=company_enumeration_text,
        company_group=company_group if company_list else "",
        optional_text=" haricindeki diğer" if company_list else "Tüm",
        num_of_APG=unique_apg_amount
    )

def format_percentage(value: float, decimal_digits: int = DEFAULT_DECIMAL_DIGITS) -> str:
    """
    Format a float value as a percentage with 'DEFAULT_DECIMAL_DIGITS' decimal points.
    """
    return "%{:.{}f}".format(value * 100, decimal_digits).replace('.', ',')

def shuffle_columns(df: pd.DataFrame, company_list: list[str]) -> pd.DataFrame:
    """
    Shuffle the columns of a DataFrame, excluding the first three and the last one.
    """
    fixed_columns = df.columns[:START_COL].tolist() + df.columns[END_COL:].tolist()
    columns_to_shuffle = [col for col in df.columns[START_COL:END_COL] if col not in company_list]
    shuffled_columns = np.random.permutation(columns_to_shuffle)
    new_column_order = fixed_columns[:START_COL] + company_list + shuffled_columns.tolist() + fixed_columns[START_COL:]
    group_df = df[new_column_order]
    group_df.columns.values[START_COL:END_COL] = COMPANIES_RANGE
    return group_df

def standardgraph(row: pd.Series) -> plt.Figure:
    """
    Create a standard scatter plot for the given row.
    """
    ax = plt.figure()
    ax = sns.scatterplot(
        data=transposed,
        x="companies",
        y=row['APG No'],
        legend=False,
        hue=company_color_indicator
    )
    ax.set(xlabel=None)
    ax.set(ylabel=None)
    ax.grid(axis='y')
    ax.set_xticks(COMPANIES_RANGE)
    ax.set_xticklabels(COMPANIES_RANGE)
    wrapped_title = wrap_text_over_words(row['APG Full Name'], WORD_WRAP_LIMIT)
    ax.set(title=wrapped_title)
    for company_index in COMPANIES_RANGE:
        value = row[company_index]
        formatted_value = format_percentage(value) if row["Birim"] == "%" else str(value)
        plt.annotate(
            text=formatted_value,
            xy=(company_index, value),
            fontsize=ANNOTATION_FONT_SIZE,
            xytext=ANNOTATION_OFFSET_PIXELS,
            textcoords="offset pixels"
        )
    if row["Birim"] == "%":
        ax.set_yticks(ax.get_yticks())
        ax.set_yticklabels(map(format_percentage, ax.get_yticks()))
    else:
        ax.set(ylabel=row["Birim"])
    ax.axhline(
        y=row["filtered_mean_iqr"],
        color=HORIZONTAL_MEAN_COLOR,
        alpha=HORIZONTAL_MEAN_ALPHA,
    )
    return ax.get_figure()

def stackedgraph(row: pd.Series) -> plt.Figure:
    """
    Create a stacked bar plot for the given row.
    """
    stacked_apg_nos = category_to_apg_dict[row["Category No"]]
    legend_labels = category_to_apg_full_name_dict[row["Category No"]]
    fig, ax = plt.subplots()
    transposed[stacked_apg_nos].plot(kind='bar', stacked=True, ax=ax)
    total_value = transposed[stacked_apg_nos].sum(axis=1).max()
    ax.legend(labels=legend_labels, loc='center left', bbox_to_anchor=(1, 0.5))
    ax.set_xticks(COMPANIES_RANGE - 1)
    ax.set_xticklabels(COMPANIES_RANGE)
    ax.set_yticks(ax.get_yticks())
    ax.set_yticklabels(map(format_percentage, ax.get_yticks()))
    if any(apg in ['E4.1', 'E4.2', 'E4.3', 'E4.4', 'E4.5'] for apg in stacked_apg_nos):
        ax.set_ylim(0, total_value)
    else:
        ax.set_ylim(0.9, total_value)
        yticks = [0.9, 1]
        ax.set_yticks(yticks)
        ax.set_yticklabels([format_percentage(tick) for tick in yticks])
    return fig

def overlayedgraph(row: pd.Series) -> plt.Figure:
    """
    Create an overlayed bar and scatter plot for the given row.
    """
    apg = row["APG Group"]
    alt_bilgi = f"{apg} EK"
    ara_df = shuffled_df[shuffled_df["APG Group"] == apg]
    main_row = ara_df[ara_df["APG No"] == apg].iloc[0]
    ek_row = ara_df[ara_df["APG No"] == alt_bilgi].iloc[0]
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.set_xticks(COMPANIES_RANGE)
    ax1.set_xticklabels(COMPANIES_RANGE)
    ax1.bar(
        x=transposed.companies,
        height=transposed[alt_bilgi],
        color=OVERLAY_GRAPH_BAR_COLOR,
        width=0.4,
        zorder=2
    )
    ax1.grid(visible=False)
    ax1.set(ylabel=f"{ek_row['APG İsmi']}\n(Çubuk Gösterim)")
    ax1.set_yticks(ax1.get_yticks())
    ax1.set_yticklabels(map(format_percentage if ek_row["Birim"] == "%" else str, ax1.get_yticks()))
    for x, y in zip(transposed.companies, transposed[alt_bilgi]):
        formatted_text = format_percentage(y) if ek_row["Birim"] == "%" else str(y)
        plt.text(
            x,
            0,
            formatted_text,
            horizontalalignment='center',
            verticalalignment='bottom'
        )
    ax2 = ax1.twinx()
    ax2.scatter(
        x=transposed.companies,
        y=transposed[main_row['APG No']],
        c=company_color_indicator,
        cmap=OVERLAY_GRAPH_COLOR_MAP,
        zorder=3
    )
    for company_index in COMPANIES_RANGE:
        value = main_row[company_index]
        formatted_value = format_percentage(value) if main_row["Birim"] == "%" else str(value)
        plt.annotate(
            text=formatted_value,
            xy=(company_index, value),
            xytext=(2,5),
            textcoords="offset pixels",
            fontsize=ANNOTATION_FONT_SIZE,
            fontweight='normal',
            zorder=4
        )
    ax2.tick_params(axis='y', labelsize=ANNOTATION_FONT_SIZE)
    ax2.grid(axis='y')
    ax2.set(ylabel=f"{main_row['APG İsmi']}\n(Nokta Gösterim)")
    ax2.set_yticks(ax2.get_yticks())
    ax2.set_yticklabels(map(format_percentage if main_row["Birim"] == "%" else str, ax2.get_yticks()))
    ax2.axhline(
        y=main_row["filtered_mean_iqr"],
        color=HORIZONTAL_MEAN_COLOR,
        alpha=HORIZONTAL_MEAN_ALPHA,
    )
    return fig
def find_consensus_outliers(outliers_iqr: list, outliers_mad: list) -> dict:
    """
    Find outliers that are detected by both IQR and MAD methods (consensus outliers).
    
    Parameters:
    outliers_iqr (list): List of tuples (column, value) from IQR method
    outliers_mad (list): List of tuples (column, value) from MAD method
    
    Returns:
    dict: Contains consensus_outliers list and consensus_count
    """
    # Extract column names from outlier lists
    iqr_columns = {col for col, _ in outliers_iqr}
    mad_columns = {col for col, _ in outliers_mad}
    
    # Find intersection (columns that appear in both methods)
    consensus_columns = iqr_columns.intersection(mad_columns)
    
    # Build consensus outlier list with values
    consensus_outliers = []
    for col, value in outliers_iqr:
        if col in consensus_columns:
            consensus_outliers.append((col, value))
    
    return {
        'consensus_outliers': consensus_outliers,
        'consensus_count': len(consensus_outliers)
    }
# Updated create_powerpoint() function section for Bulgu text
# Updated create_powerpoint() function section for Bulgu text
def create_powerpoint():
    """
    Create a PowerPoint presentation with the graphs for the given company group.
    """
    graphics_save_directory = GRAPHICS_DIRECTORY / f'{REPORT_YEAR}_{REPORT_TYPE}' / company_group
    graphics_save_directory.mkdir(parents=True, exist_ok=True)
    presentation = Presentation(TEMPLATE_PATH)
    bulgu_shapes = [
        shape for slide in presentation.slides
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and shape.text.startswith("Bulgu")
    ]
    if len(bulgu_shapes) != shuffled_df['Bulgu?'].sum():
        raise SystemExit(f"The number of 'Bulgu' shapes in the presentation does not match the number of 'Bulgu' rows in the DataFrame. Check the 'Bulgu?' column in the Excel file.")
    bulgu_iterator = iter(bulgu_shapes)
    
    for _, row in shuffled_df.iterrows():
        grafik_tipi = row["Grafik_tipi"]
        if grafik_tipi == "standard":
            fig = standardgraph(row)
        elif grafik_tipi == "stacked":
            fig = stackedgraph(row)
        elif grafik_tipi == "overlayed":
            fig = overlayedgraph(row)
        else:
            raise SystemExit(f"There is no function that corresponds to the {grafik_tipi} graph type. Please check the 'Grafik_tipi' column in the Excel file.")
        
        apg_pic_name = f'{row["APG Full Name"]}.png'
        clean_apg_path = ILLEGAL_WINDOWS_PATH_CHARACTERS.sub('_', apg_pic_name)
        pic_path = graphics_save_directory / clean_apg_path
        pic_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(pic_path, bbox_inches='tight')
        logger.info(f"Saved the figure for {row['APG No']} to {pic_path.relative_to(MAIN_DIRECTORY)}")
        plt.close(fig)
        
        slide = presentation.slides[row["Sayfa"] - 1]
        left, top, height, width = row["Left"], row["Top"], row["Height"], row["Width"]
        slide.shapes.add_picture(str(pic_path), left, top, width, height)
        
        if row["Bulgu?"]:
            shape = next(bulgu_iterator)
            def format_value(value, birim):
                if pd.isna(value):
                    return "N/A"
                return format_percentage(value) if birim == "%" else round(value, DEFAULT_DECIMAL_DIGITS)
            
            iqr_mean = format_value(row["filtered_mean_iqr"], row["Birim"])
            mad_mean = format_value(row["filtered_mean_mad"], row["Birim"])
            consensus_count = row["consensus_outlier_count"]
            
            # NEW: Updated Bulgu text to include consensus outliers
            bulgu_text = (
                f"Bulgu\n{NUM_OF_COMPANIES} şirketin ortalaması:\n"
                f"- Method 1: {iqr_mean} (Aykırı: {row['outlier_count_iqr']})\n"
                f"- Method 2: {mad_mean} (Aykırı: {row['outlier_count_mad']})\n"
                f"- Kesişim Aykırı Değer: {consensus_count} şirket"  # New consensus line
            )
            
            shape.text = bulgu_text
            paragraphs = shape.text_frame.paragraphs
            paragraphs[0].font.bold = True
            for paragraph in paragraphs:
                paragraph.font.size = Pt(FONT_SIZE)
    
    ay = 6 if REPORT_TYPE == REPORT_TYPE_CHOICES[0] else 12
    presentation.slides[0].shapes[3].text = f'{REPORT_YEAR} Yılı {ay} Aylık Döneme Ait Performans Göstergesi Sonuçları'
    presentation.slides[1].shapes[3].text = generate_presentation_intro_text(company_list)
    presentation_path = REPORTS_DIRECTORY / f'{REPORT_YEAR}_{REPORT_TYPE}' / company_group / f'Kıyaslama Raporu {REPORT_YEAR}_{REPORT_TYPE}_{company_group}.pptx'
    presentation_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(presentation_path)
    logger.info(f"Saved the presentation for {company_group} to {presentation_path.relative_to(MAIN_DIRECTORY)}")

if __name__ == "__main__":
    dataframe_dict = pd.read_excel(MASTER_FILE, sheet_name=[f"{REPORT_YEAR}_Total_Veriler", "pptx_layout"])
    df = dataframe_dict[f"{REPORT_YEAR}_Total_Veriler"]
    pptx_layout = dataframe_dict["pptx_layout"]
    merged_df = pd.merge(df, pptx_layout, left_on='APG No', right_on='APG Kodu', how='left')
    merged_df['Sayfa'] = merged_df['Sayfa'].fillna(0).astype(int)
    merged_df['Category No'] = merged_df['APG No'].str.split('.').str[0]
    merged_df['APG Full Name'] = merged_df.apply(lambda row: f'{row["APG No"]}-{row["APG İsmi"]}', axis=1)
    merged_df['APG Group'] = merged_df['APG No'].str.extract(APG_NO_PATTERN)[0]
    unique_apg_amount = merged_df['APG Group'].nunique()
    stacked_df = merged_df[merged_df.Grafik_tipi == "stacked"][["Category No", "APG No", "APG Full Name"]]
    category_to_apg_dict = stacked_df.groupby('Category No')['APG No'].apply(list).to_dict()
    category_to_apg_full_name_dict = stacked_df.groupby('Category No')['APG Full Name'].apply(list).to_dict()
    for company_group, company_list in COMPANY_GROUPS.items():
        if company_group in COMPANY_GROUPS_EXCLUDED_FROM_REPORT:
            continue
        num_of_group_companies = len(company_list)
        num_of_other_companies = NUM_OF_COMPANIES - num_of_group_companies
        company_color_indicator = [GROUP_COMPANY_INDICATOR] * num_of_group_companies + [RIVAL_COMPANY_INDICATOR] * num_of_other_companies
        shuffled_groups = [
            shuffle_columns(group, company_list)
            for _, group
            in merged_df.groupby('Category No', sort=False)
        ]
        shuffled_df = pd.concat(shuffled_groups).reset_index(drop=True)
        logger.info(f"Processing outlier detection for company group: {company_group}")
        std_results = shuffled_df.apply(
            lambda row: filtered_mean_with_outliers(row, START_COL, END_COL, SIGMA),
            axis=1
        )
        shuffled_df['filtered_mean_std'] = std_results.apply(lambda x: x['filtered_mean'])
        shuffled_df['outliers_std'] = std_results.apply(lambda x: x['outliers'])
        shuffled_df['outlier_count_std'] = std_results.apply(lambda x: x['outlier_count'])
        iqr_results = shuffled_df.apply(
            lambda row: filtered_mean_with_outliers_iqr(row, START_COL, END_COL, IQR_FACTOR),
            axis=1
        )
        shuffled_df['filtered_mean_iqr'] = iqr_results.apply(lambda x: x['filtered_mean'])
        shuffled_df['outliers_iqr'] = iqr_results.apply(lambda x: x['outliers'])
        shuffled_df['outlier_count_iqr'] = iqr_results.apply(lambda x: x['outlier_count'])
        mad_results = shuffled_df.apply(
            lambda row: filtered_mean_with_outliers_mad(row, START_COL, END_COL, MAD_THRESHOLD),
            axis=1
        )
        shuffled_df['filtered_mean_mad'] = mad_results.apply(lambda x: x['filtered_mean'])
        shuffled_df['outliers_mad'] = mad_results.apply(lambda x: x['outliers'])
        shuffled_df['outlier_count_mad'] = mad_results.apply(lambda x: x['outlier_count'])

        # NEW: Consensus outlier detection (IQR + MAD agreement)
        consensus_results = shuffled_df.apply(
            lambda row: find_consensus_outliers(
                row['outliers_iqr'], 
                row['outliers_mad']
            ),
            axis=1
        )
        shuffled_df['consensus_outliers'] = consensus_results.apply(lambda x: x['consensus_outliers'])
        shuffled_df['consensus_outlier_count'] = consensus_results.apply(lambda x: x['consensus_count'])
        
        # Updated report columns to include consensus outliers
        report_columns = [
            "APG No", "APG İsmi", "Birim", *COMPANIES_RANGE,
            "filtered_mean_std", "outlier_count_std", "outliers_std",
            "filtered_mean_iqr", "outlier_count_iqr", "outliers_iqr",
            "filtered_mean_mad", "outlier_count_mad", "outliers_mad",
            "consensus_outlier_count", "consensus_outliers"  # New consensus columns
        ]
        
        report_df = shuffled_df[report_columns]
        REPORT_XLSX_PATH = REPORTS_DIRECTORY / f'{REPORT_YEAR}_{REPORT_TYPE}' / company_group / f"{REPORT_YEAR}_{REPORT_TYPE}_{company_group}_Shuffled.xlsx"
        report_df.to_excel(REPORT_XLSX_PATH, index=False)
        logger.info(f"Saved shuffled DataFrame to {REPORT_XLSX_PATH.relative_to(MAIN_DIRECTORY)}")
        transposable = shuffled_df[["APG No"] + list(COMPANIES_RANGE)]
        transposed = transposable.set_index("APG No").T.reset_index().rename(columns={"index": "companies"})
        create_powerpoint()
