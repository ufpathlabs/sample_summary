#!/usr/bin/env python3

import argparse
import os
import pandas as pd
from functools import reduce
import xlsxwriter

def parse_sample_id_args():
    """
    """
    parser = argparse.ArgumentParser(
        description = 'sample id for json data files and qc files'
    )
    parser.add_argument(
        '-p',
        metavar = '--PATH',
        type = str,
        help = 'GS700-2211-NQ22-0014',
        required = True
    )
    args = parser.parse_args()
    return args


def get_summary_data(directory, run_id):
    """
    """
    df_list = []
    for sample_dir in os.listdir(directory):
        if f'{run_id}_BC' in sample_dir:
            sample_info = sample_dir.split('_')
            sample_id = f'{sample_info[0]}_{sample_info[1]}_{sample_info[2]}'
            summary_csv = f'{directory}/{sample_dir}/Additional Files/{sample_id}.summary.csv'
            coverage_csv = f'{directory}/{sample_dir}/{sample_id}.qc-coverage-region-1_coverage_metrics.csv'
            df = pd.read_csv(
                summary_csv,
                skiprows=4,
                names = ['DRAGEN Enrichment Summary Report', f'{sample_id}_Value']
            )
            cov_df = pd.read_csv(
                coverage_csv,
                usecols=[2,3],
                names = ['DRAGEN Enrichment Summary Report', f'{sample_id}_Value']
            )
            entry = cov_df.loc[cov_df['DRAGEN Enrichment Summary Report'] == 'PCT of QC coverage region with coverage [100x: inf)']
            new_df = pd.concat([df, entry])
            df_list.append(new_df)
    return df_list


def merge_summary_data(df_list):
    """
    """
    return reduce(lambda  left,right: pd.merge(left,right,on=['DRAGEN Enrichment Summary Report'],how='outer'), df_list)


def write_run_summary_xlsx(directory, run_id, df_merged):
    """
    """
    with pd.ExcelWriter(f'{directory}/{run_id}_Sample Metrics.xlsx') as writer:
        df_merged.to_excel(writer, sheet_name = f'{run_id}', index = False, na_rep = 'NA')
        for column in df_merged:
            col_idx = df_merged.columns.get_loc(column)
            writer.sheets[f'{run_id}'].set_column(col_idx, col_idx, 39)

        workbook = writer.book
        worksheet = writer.sheets[f'{run_id}']
        format1 = workbook.add_format({'bg_color': '#CCEECE', 'font_color': '#225F00'})
        border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})

        for r in range(0,len(df_merged.index)):
            if df_merged.iat[r,0] == "Mean target coverage depth":
                worksheet.set_row(r+1, None, format1)
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df_merged), len(df_merged.columns)-1), {'type': 'no_errors', 'format': border_fmt})


def main():
    """
    """
    ASSAY_DIR = "/ext/path/DRL/Molecular/NGS21/ASSAYS"
    args = parse_sample_id_args()
    assay_info = args.p.split('-')
    assay = assay_info[0]
    run_info = assay_info[2] + assay_info[3]
    run_id = f'{run_info[0:2]}-{run_info[2:4]}-{run_info[6:8]}'
    directory = f'{ASSAY_DIR}/{assay}/{args.p}/'
    df_list = get_summary_data(directory, run_id)
    df_merged = merge_summary_data(df_list)
    write_run_summary_xlsx(directory, run_id, df_merged)

if __name__ == '__main__':
    main()
