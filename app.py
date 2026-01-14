import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import numpy as np
import pyomo.environ as pyo
from pyomo.opt import SolverFactory, TerminationCondition
from pathlib import Path
import json
from datetime import datetime
import tempfile
import io

# ================================================================
# PAGE CONFIGURATION
# ================================================================
st.set_page_config(
    page_title="MCDM Criteria Selection Tool",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================================================================
# SESSION STATE INITIALIZATION
# ================================================================
if 'data' not in st.session_state:
    st.session_state.data = None
if 'weights' not in st.session_state:
    st.session_state.weights = None
if 'model' not in st.session_state:
    st.session_state.model = None
if 'result' not in st.session_state:
    st.session_state.result = None
if 'config' not in st.session_state:
    st.session_state.config = None

# ================================================================
# EXCEL TEMPLATE GENERATOR - COMPLETE VERSION
# ================================================================

def generate_excel_template(num_criteria, num_alternatives, num_experts, num_objectives,
                           omega, zeta, alpha, gamma_O, gamma_S, delta, theta,
                           tau_O, tau_S, lambda_th, mu):
    """Generate complete Excel template with all 11 sheets"""
    
    # Store config
    st.session_state.config = {
        'num_criteria': num_criteria,
        'num_alternatives': num_alternatives,
        'num_experts': num_experts,
        'num_objectives': num_objectives,
        'omega': omega,
        'zeta': zeta,
        'alpha': alpha,
        'gamma_O': gamma_O,
        'gamma_S': gamma_S,
        'delta': delta,
        'theta': theta,
        'tau_O': tau_O,
        'tau_S': tau_S,
        'lambda': lambda_th,
        'mu': mu
    }
    
    # Calculate row positions
    CRITERIA_START_ROW = 11
    ALTERNATIVES_START_ROW = 11 + num_criteria + 3
    OBJECTIVES_START_ROW = ALTERNATIVES_START_ROW + num_alternatives + 3
    
    # Create workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # Define styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    output_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    section_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # ================================================================
    # SHEET 0: CONFIGURATION
    # ================================================================
    ws_config = wb.create_sheet("0_Configuration")
    ws_config['A1'] = "MCDM CRITERIA SELECTION - CONFIGURATION"
    ws_config['A1'].font = Font(bold=True, size=14)
    ws_config.merge_cells('A1:D1')
    
    row = 3
    ws_config[f'A{row}'] = "PROBLEM STRUCTURE"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    structure_data = [
        ["Number of Criteria", num_criteria],
        ["Number of Alternatives", num_alternatives],
        ["Number of Experts", num_experts],
        ["Number of Objectives", num_objectives],
    ]
    
    for label, value in structure_data:
        ws_config[f'A{row}'] = label
        ws_config[f'B{row}'] = value
        row += 1
    
    row += 1
    
    # CRITERIA DEFINITION TABLE
    ws_config[f'A{row}'] = "CRITERIA DEFINITIONS (Fill in the yellow cells)"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    headers = ["Criterion ID", "Criterion Name", "Type (Cost/Benefit)", "Description (Optional)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws_config.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    row += 1
    
    for i in range(num_criteria):
        ws_config.cell(row=row, column=1, value=f"C{i+1}")
        
        name_cell = ws_config.cell(row=row, column=2, value=f"Criterion {i+1}")
        name_cell.fill = input_fill
        name_cell.border = thin_border
        
        type_cell = ws_config.cell(row=row, column=3, value="Benefit")
        type_cell.fill = input_fill
        type_cell.border = thin_border
        
        desc_cell = ws_config.cell(row=row, column=4, value="")
        desc_cell.fill = input_fill
        desc_cell.border = thin_border
        
        row += 1
    
    dv = DataValidation(type="list", formula1='"Cost,Benefit"', allow_blank=False)
    ws_config.add_data_validation(dv)
    type_range = f"C{CRITERIA_START_ROW}:C{CRITERIA_START_ROW + num_criteria - 1}"
    dv.add(type_range)
    
    row += 1
    
    # ALTERNATIVES
    ws_config[f'A{row}'] = "ALTERNATIVES DEFINITIONS (Fill in the yellow cells)"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    headers = ["Alternative ID", "Alternative Name", "Description (Optional)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws_config.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    row += 1
    
    for i in range(num_alternatives):
        ws_config.cell(row=row, column=1, value=f"A{i+1}")
        
        name_cell = ws_config.cell(row=row, column=2, value=f"Alternative {i+1}")
        name_cell.fill = input_fill
        name_cell.border = thin_border
        
        desc_cell = ws_config.cell(row=row, column=3, value="")
        desc_cell.fill = input_fill
        desc_cell.border = thin_border
        
        row += 1
    
    row += 1
    
    # OBJECTIVES
    ws_config[f'A{row}'] = "OBJECTIVES DEFINITIONS (Fill in the yellow cells)"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    headers = ["Objective ID", "Objective Name", "Description (Optional)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws_config.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    row += 1
    
    for i in range(num_objectives):
        ws_config.cell(row=row, column=1, value=f"O{i+1}")
        
        name_cell = ws_config.cell(row=row, column=2, value=f"Objective {i+1}")
        name_cell.fill = input_fill
        name_cell.border = thin_border
        
        desc_cell = ws_config.cell(row=row, column=3, value="")
        desc_cell.fill = input_fill
        desc_cell.border = thin_border
        
        row += 1
    
    row += 2
    
    # PARSIMONY
    ws_config[f'A{row}'] = "PARSIMONY BOUNDS (Step 5)"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    parsimony_data = [
        ["Target Minimum (ω)", omega],
        ["Target Maximum (ζ)", zeta],
    ]
    
    for label, value in parsimony_data:
        ws_config[f'A{row}'] = label
        ws_config[f'B{row}'] = value
        row += 1
    
    row += 1
    
    # THRESHOLDS
    ws_config[f'A{row}'] = "THRESHOLDS"
    ws_config[f'A{row}'].font = Font(bold=True, size=12)
    ws_config[f'A{row}'].fill = section_fill
    ws_config.merge_cells(f'A{row}:D{row}')
    row += 1
    
    threshold_data = [
        ["Step 1: Completeness (α)", alpha],
        ["Step 3: Measurability Objective (γ_O)", gamma_O],
        ["Step 3: Measurability Subjective (γ_S)", gamma_S],
        ["Step 4: Distinctiveness (δ)", delta],
        ["Step 6: Sensitivity (θ)", theta],
        ["Step 7: Cost-effectiveness Objective (τ_O)", tau_O],
        ["Step 7: Cost-effectiveness Subjective (τ_S)", tau_S],
        ["Step 8: Alignment (λ)", lambda_th],
        ["Step 9: Cognitive Coherence (μ)", mu],
    ]
    
    for label, value in threshold_data:
        ws_config[f'A{row}'] = label
        ws_config[f'B{row}'] = value
        row += 1
    
    ws_config.column_dimensions['A'].width = 40
    ws_config.column_dimensions['B'].width = 20
    ws_config.column_dimensions['C'].width = 20
    ws_config.column_dimensions['D'].width = 30
    
    # ================================================================
    # SHEET 1: COMPLETENESS
    # ================================================================
    ws1 = wb.create_sheet("1_Completeness")
    ws1['A1'] = "Step 1: Completeness Evaluation"
    ws1['A1'].font = Font(bold=True, size=12)
    ws1['A2'] = f"Rate how well each criterion covers the decision aspect (1-10 scale). Threshold: α = {alpha}"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["Median", "Status"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws1.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws1.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws1.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws1.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        median_col = get_column_letter(3 + num_experts)
        
        median_cell = ws1.cell(row=row_num, column=3 + num_experts)
        median_cell.value = f'=MEDIAN({first_col}{row_num}:{last_col}{row_num})'
        median_cell.fill = output_fill
        median_cell.border = thin_border
        median_cell.number_format = '0.00'
        
        status_cell = ws1.cell(row=row_num, column=4 + num_experts)
        status_cell.value = f'=IF({median_col}{row_num}>={alpha},"Meets","Below")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
    
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 30
    for e in range(num_experts + 2):
        ws1.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 2: OBJECTIVITY
    # ================================================================
    ws2 = wb.create_sheet("2_Objectivity")
    ws2['A1'] = "Step 2: Objectivity/Subjectivity Classification"
    ws2['A1'].font = Font(bold=True, size=12)
    ws2['A2'] = "Classify each criterion: 1 = Objective, 0 = Subjective (Majority vote determines final classification)"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["Sum", "Final Class", "Binary"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws2.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws2.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws2.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws2.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        sum_col = get_column_letter(3 + num_experts)
        
        sum_cell = ws2.cell(row=row_num, column=3 + num_experts)
        sum_cell.value = f'=SUM({first_col}{row_num}:{last_col}{row_num})'
        sum_cell.fill = output_fill
        sum_cell.border = thin_border
        
        class_col = get_column_letter(4 + num_experts)
        class_cell = ws2.cell(row=row_num, column=4 + num_experts)
        class_cell.value = f'=IF({sum_col}{row_num}>{num_experts}/2,"Objective","Subjective")'
        class_cell.fill = output_fill
        class_cell.border = thin_border
        
        binary_cell = ws2.cell(row=row_num, column=5 + num_experts)
        binary_cell.value = f'=IF({class_col}{row_num}="Objective",1,0)'
        binary_cell.fill = output_fill
        binary_cell.border = thin_border
    
    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 30
    for e in range(num_experts + 3):
        ws2.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 3: MEASURABILITY
    # ================================================================
    ws3 = wb.create_sheet("3_Measurability")
    ws3['A1'] = "Step 3: Measurability Assessment"
    ws3['A1'].font = Font(bold=True, size=12)
    ws3['A2'] = f"Rate how easily each criterion can be quantified (1-10 scale). Thresholds: γ_O = {gamma_O}, γ_S = {gamma_S}"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["Median", "Type", "Threshold γ_i", "Status"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws3.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws3.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws3.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws3.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        median_col = get_column_letter(3 + num_experts)
        
        median_cell = ws3.cell(row=row_num, column=3 + num_experts)
        median_cell.value = f'=MEDIAN({first_col}{row_num}:{last_col}{row_num})'
        median_cell.fill = output_fill
        median_cell.border = thin_border
        median_cell.number_format = '0.00'
        
        type_col = get_column_letter(4 + num_experts)
        type_cell = ws3.cell(row=row_num, column=4 + num_experts)
        type_cell.value = f'=2_Objectivity!$H${5 + i}'
        type_cell.fill = output_fill
        type_cell.border = thin_border
        
        thresh_col = get_column_letter(5 + num_experts)
        thresh_cell = ws3.cell(row=row_num, column=5 + num_experts)
        thresh_cell.value = f'=IF({type_col}{row_num}=1,{gamma_O},{gamma_S})'
        thresh_cell.fill = output_fill
        thresh_cell.border = thin_border
        thresh_cell.number_format = '0.00'
        
        status_cell = ws3.cell(row=row_num, column=6 + num_experts)
        status_cell.value = f'=IF({median_col}{row_num}>={thresh_col}{row_num},"Meets","Below")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
    
    ws3.column_dimensions['A'].width = 12
    ws3.column_dimensions['B'].width = 30
    for e in range(num_experts + 4):
        ws3.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 4: DISTINCTIVENESS
    # ================================================================
    ws4 = wb.create_sheet("4_Distinctiveness")
    ws4['A1'] = "Step 4: Distinctiveness - Decision Matrices"
    ws4['A1'].font = Font(bold=True, size=12)
    ws4['A2'] = f"Provide decision matrices for each expert. Correlation threshold: δ = {delta}"
    ws4['A3'] = "Note: Correlation analysis will be performed externally in Python"
    
    row = 5
    for e in range(num_experts):
        ws4.cell(row=row, column=1, value=f"Expert {e+1} Decision Matrix").font = Font(bold=True)
        row += 1
        
        headers = ["Alternative"]
        for c in range(num_criteria):
            headers.append(f"C{c+1}")
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws4.cell(row=row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
        row += 1
        
        for a in range(num_alternatives):
            alt_cell = ws4.cell(row=row, column=1)
            alt_cell.value = f'=0_Configuration!$B${ALTERNATIVES_START_ROW + 1 + a}'
            alt_cell.border = thin_border
            
            for c in range(num_criteria):
                cell = ws4.cell(row=row, column=2+c)
                cell.fill = input_fill
                cell.border = thin_border
            row += 1
        
        row += 2
    
    ws4.column_dimensions['A'].width = 35
    for c in range(num_criteria):
        ws4.column_dimensions[get_column_letter(2+c)].width = 10
    
    # ================================================================
    # SHEET 6: SENSITIVITY
    # ================================================================
    ws6 = wb.create_sheet("6_Sensitivity")
    ws6['A1'] = "Step 6: Sensitivity Analysis - Decision Matrices"
    ws6['A1'].font = Font(bold=True, size=12)
    ws6['A2'] = f"Provide decision matrices for each expert. Elasticity threshold: θ = {theta}"
    ws6['A3'] = "Note: Sensitivity analysis will be performed externally in Python"
    
    row = 5
    for e in range(num_experts):
        ws6.cell(row=row, column=1, value=f"Expert {e+1} Decision Matrix").font = Font(bold=True)
        row += 1
        
        headers = ["Alternative"]
        for c in range(num_criteria):
            headers.append(f"C{c+1}")
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws6.cell(row=row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
        row += 1
        
        for a in range(num_alternatives):
            alt_cell = ws6.cell(row=row, column=1)
            alt_cell.value = f'=0_Configuration!$B${ALTERNATIVES_START_ROW + 1 + a}'
            alt_cell.border = thin_border
            
            for c in range(num_criteria):
                cell = ws6.cell(row=row, column=2+c)
                cell.fill = input_fill
                cell.border = thin_border
            row += 1
        
        row += 2
    
    ws6.column_dimensions['A'].width = 35
    for c in range(num_criteria):
        ws6.column_dimensions[get_column_letter(2+c)].width = 10
    
    # ================================================================
    # SHEET 7: COST-EFFECTIVENESS
    # ================================================================
    ws7 = wb.create_sheet("7_Cost_Effectiveness")
    ws7['A1'] = "Step 7: Cost-Effectiveness Evaluation"
    ws7['A1'].font = Font(bold=True, size=12)
    ws7['A2'] = f"Rate cost-effectiveness (0-10 Likert scale). Thresholds: τ_O = {tau_O}, τ_S = {tau_S}"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["Median", "Type", "Threshold τ_i", "Status", "Binary"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws7.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws7.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws7.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws7.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        median_col = get_column_letter(3 + num_experts)
        
        median_cell = ws7.cell(row=row_num, column=3 + num_experts)
        median_cell.value = f'=MEDIAN({first_col}{row_num}:{last_col}{row_num})'
        median_cell.fill = output_fill
        median_cell.border = thin_border
        median_cell.number_format = '0.00'
        
        type_col = get_column_letter(4 + num_experts)
        type_cell = ws7.cell(row=row_num, column=4 + num_experts)
        type_cell.value = f'=2_Objectivity!$H${5 + i}'
        type_cell.fill = output_fill
        type_cell.border = thin_border
        
        thresh_col = get_column_letter(5 + num_experts)
        thresh_cell = ws7.cell(row=row_num, column=5 + num_experts)
        thresh_cell.value = f'=IF({type_col}{row_num}=1,{tau_O},{tau_S})'
        thresh_cell.fill = output_fill
        thresh_cell.border = thin_border
        thresh_cell.number_format = '0.00'
        
        status_col = get_column_letter(6 + num_experts)
        status_cell = ws7.cell(row=row_num, column=6 + num_experts)
        status_cell.value = f'=IF({median_col}{row_num}>={thresh_col}{row_num},"Meets","Below")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
        
        binary_cell = ws7.cell(row=row_num, column=7 + num_experts)
        binary_cell.value = f'=IF({status_col}{row_num}="Meets",1,0)'
        binary_cell.fill = output_fill
        binary_cell.border = thin_border
    
    ws7.column_dimensions['A'].width = 12
    ws7.column_dimensions['B'].width = 30
    for e in range(num_experts + 5):
        ws7.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 8: ALIGNMENT
    # ================================================================
    ws8 = wb.create_sheet("8_Alignment")
    ws8['A1'] = "Step 8: Alignment Assessment"
    ws8['A1'].font = Font(bold=True, size=12)
    ws8['A2'] = f"Rate criterion-objective alignment (1-10 scale). Threshold: λ = {lambda_th}"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["Median", "Status"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws8.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws8.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws8.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws8.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        median_col = get_column_letter(3 + num_experts)
        
        median_cell = ws8.cell(row=row_num, column=3 + num_experts)
        median_cell.value = f'=MEDIAN({first_col}{row_num}:{last_col}{row_num})'
        median_cell.fill = output_fill
        median_cell.border = thin_border
        median_cell.number_format = '0.00'
        
        status_cell = ws8.cell(row=row_num, column=4 + num_experts)
        status_cell.value = f'=IF({median_col}{row_num}>={lambda_th},"Meets","Below")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
    
    ws8.column_dimensions['A'].width = 12
    ws8.column_dimensions['B'].width = 30
    for e in range(num_experts + 2):
        ws8.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 9: COGNITIVE COHERENCE
    # ================================================================
    ws9 = wb.create_sheet("9_Cognitive_Coherence")
    ws9['A1'] = "Step 9: Cognitive Coherence"
    ws9['A1'].font = Font(bold=True, size=12)
    ws9['A2'] = f"Cross-expert ratings of definitions (no self-ratings). Threshold: μ = {mu}"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for rater in range(num_experts):
        for author in range(num_experts):
            if rater != author:
                headers.append(f"E{rater+1}→E{author+1}")
    headers.extend(["Median", "Status"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws9.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border
    
    num_cross_ratings = num_experts * (num_experts - 1)
    for i in range(num_criteria):
        row_num = 5 + i
        ws9.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws9.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for j in range(num_cross_ratings):
            cell = ws9.cell(row=row_num, column=3+j)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_cross_ratings)
        median_col = get_column_letter(3 + num_cross_ratings)
        
        median_cell = ws9.cell(row=row_num, column=3 + num_cross_ratings)
        median_cell.value = f'=MEDIAN({first_col}{row_num}:{last_col}{row_num})'
        median_cell.fill = output_fill
        median_cell.border = thin_border
        median_cell.number_format = '0.00'
        
        status_cell = ws9.cell(row=row_num, column=4 + num_cross_ratings)
        status_cell.value = f'=IF({median_col}{row_num}>={mu},"Meets","Below")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
    
    ws9.column_dimensions['A'].width = 12
    ws9.column_dimensions['B'].width = 30
    for j in range(num_cross_ratings + 2):
        ws9.column_dimensions[get_column_letter(3+j)].width = 10
    
    # ================================================================
    # SHEET 10: MONOTONE COHERENCE
    # ================================================================
    ws10 = wb.create_sheet("10_Monotone_Coherence")
    ws10['A1'] = "Step 10: Monotone Coherence"
    ws10['A1'].font = Font(bold=True, size=12)
    ws10['A2'] = "Binary votes on monotonicity (1 = monotone, 0 = not monotone)"
    
    row = 4
    headers = ["Criterion ID", "Criterion Name"]
    for e in range(num_experts):
        headers.append(f"Expert {e+1}")
    headers.extend(["q_i (unanimity)", "Status"])
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws10.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    for i in range(num_criteria):
        row_num = 5 + i
        ws10.cell(row=row_num, column=1, value=f"C{i+1}")
        
        name_cell = ws10.cell(row=row_num, column=2)
        name_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + i}'
        name_cell.border = thin_border
        
        for e in range(num_experts):
            cell = ws10.cell(row=row_num, column=3+e)
            cell.fill = input_fill
            cell.border = thin_border
        
        first_col = get_column_letter(3)
        last_col = get_column_letter(2 + num_experts)
        q_col = get_column_letter(3 + num_experts)
        
        q_cell = ws10.cell(row=row_num, column=3 + num_experts)
        q_cell.value = f'=PRODUCT({first_col}{row_num}:{last_col}{row_num})'
        q_cell.fill = output_fill
        q_cell.border = thin_border
        
        status_cell = ws10.cell(row=row_num, column=4 + num_experts)
        status_cell.value = f'=IF({q_col}{row_num}=1,"Meets","Does not meet")'
        status_cell.fill = output_fill
        status_cell.border = thin_border
    
    ws10.column_dimensions['A'].width = 12
    ws10.column_dimensions['B'].width = 30
    for e in range(num_experts + 2):
        ws10.column_dimensions[get_column_letter(3+e)].width = 12
    
    # ================================================================
    # SHEET 11: REPRESENTATIVENESS
    # ================================================================
    ws11 = wb.create_sheet("11_Representativeness")
    ws11['A1'] = "Step 11: Representativeness"
    ws11['A1'].font = Font(bold=True, size=12)
    ws11['A2'] = "Assign criteria to objectives (1 = assigned, 0 = not; max one per criterion per expert)"
    
    expert_data_rows = []
    row = 5
    
    for e in range(num_experts):
        ws11.cell(row=row, column=1, value=f"Expert {e+1} Assignments").font = Font(bold=True)
        row += 1
        
        headers = ["Criterion"]
        for o in range(num_objectives):
            headers.append(f"O{o+1}")
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws11.cell(row=row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
        row += 1
        
        expert_start_row = row
        expert_data_rows.append(expert_start_row)
        
        for c in range(num_criteria):
            crit_cell = ws11.cell(row=row, column=1)
            crit_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + c}'
            crit_cell.border = thin_border
            
            for o in range(num_objectives):
                cell = ws11.cell(row=row, column=2+o)
                cell.fill = input_fill
                cell.border = thin_border
            row += 1
        
        row += 2
    
    row += 2
    ws11.cell(row=row, column=1, value="CONSOLIDATED (Majority Vote)").font = Font(bold=True, size=12)
    row += 2
    
    headers = ["Criterion"]
    for o in range(num_objectives):
        headers.append(f"O{o+1}")
    headers.append("e_i^{rp}")
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws11.cell(row=row, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    consolidated_row_start = row + 1
    
    for c in range(num_criteria):
        row += 1
        crit_cell = ws11.cell(row=row, column=1)
        crit_cell.value = f'=0_Configuration!$B${CRITERIA_START_ROW + c}'
        crit_cell.border = thin_border
        
        for o in range(num_objectives):
            obj_col = get_column_letter(2 + o)
            
            vote_refs = []
            for e in range(num_experts):
                expert_row = expert_data_rows[e] + c
                vote_refs.append(f"{obj_col}{expert_row}")
            
            sum_formula = "+".join(vote_refs)
            majority_formula = f'=IF({sum_formula}>{num_experts}/2,1,0)'
            
            cell = ws11.cell(row=row, column=2+o)
            cell.value = majority_formula
            cell.fill = output_fill
            cell.border = thin_border
        
        first_obj_col = get_column_letter(2)
        last_obj_col = get_column_letter(1 + num_objectives)
        
        e_rp_cell = ws11.cell(row=row, column=2 + num_objectives)
        e_rp_cell.value = f'=MIN(1,SUM({first_obj_col}{row}:{last_obj_col}{row}))'
        e_rp_cell.fill = output_fill
        e_rp_cell.border = thin_border
    
    ws11.column_dimensions['A'].width = 35
    for o in range(num_objectives + 1):
        ws11.column_dimensions[get_column_letter(2+o)].width = 10
    
    # Save to BytesIO
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer


# ================================================================
# EXCEL READER - COMPLETE VERSION
# ================================================================

def read_mcdm_template(file):
    """Read filled MCDM Excel template and extract all data"""
    
    results = {}
    
    # Read configuration
    df_config = pd.read_excel(file, sheet_name='0_Configuration', header=None)
    
    def find_row_with_text(df, text):
        for idx, row in df.iterrows():
            if pd.notna(row[0]) and text.upper() in str(row[0]).upper():
                return idx
        return None
    
    num_criteria = int(df_config.iloc[3, 1])
    num_alternatives = int(df_config.iloc[4, 1])
    num_experts = int(df_config.iloc[5, 1])
    num_objectives = int(df_config.iloc[6, 1])
    
    results['num_criteria'] = num_criteria
    results['num_alternatives'] = num_alternatives
    results['num_experts'] = num_experts
    results['num_objectives'] = num_objectives
    
    # Extract criteria
    criteria_header_row = find_row_with_text(df_config, "CRITERIA DEFINITIONS")
    criteria_start_row = criteria_header_row + 2
    criteria_ids = []
    criteria_names = []
    criteria_types = []
    
    for i in range(num_criteria):
        row_idx = criteria_start_row + i
        criteria_ids.append(str(df_config.iloc[row_idx, 0]))
        criteria_names.append(str(df_config.iloc[row_idx, 1]))
        criteria_types.append(str(df_config.iloc[row_idx, 2]))
    
    results['criteria_ids'] = criteria_ids
    results['criteria_names'] = criteria_names
    results['criteria_types'] = criteria_types
    
    # Extract alternatives
    alternatives_header_row = find_row_with_text(df_config, "ALTERNATIVES DEFINITIONS")
    alternatives_start_row = alternatives_header_row + 2
    alternatives_ids = []
    alternatives_names = []
    
    for i in range(num_alternatives):
        row_idx = alternatives_start_row + i
        alternatives_ids.append(str(df_config.iloc[row_idx, 0]))
        alternatives_names.append(str(df_config.iloc[row_idx, 1]))
    
    results['alternatives_ids'] = alternatives_ids
    results['alternatives_names'] = alternatives_names
    
    # Extract objectives
    objectives_header_row = find_row_with_text(df_config, "OBJECTIVES DEFINITIONS")
    objectives_start_row = objectives_header_row + 2
    objectives_ids = []
    objectives_names = []
    
    for i in range(num_objectives):
        row_idx = objectives_start_row + i
        objectives_ids.append(str(df_config.iloc[row_idx, 0]))
        objectives_names.append(str(df_config.iloc[row_idx, 1]))
    
    results['objectives_ids'] = objectives_ids
    results['objectives_names'] = objectives_names
    
    # Extract thresholds
    parsimony_header_row = find_row_with_text(df_config, "PARSIMONY")
    parsimony_start_row = parsimony_header_row + 1
    results['omega'] = int(df_config.iloc[parsimony_start_row, 1])
    results['zeta'] = int(df_config.iloc[parsimony_start_row + 1, 1])
    
    thresholds_header_row = find_row_with_text(df_config, "THRESHOLDS")
    thresholds_start_row = thresholds_header_row + 1
    results['alpha'] = float(df_config.iloc[thresholds_start_row, 1])
    results['gamma_O'] = float(df_config.iloc[thresholds_start_row + 1, 1])
    results['gamma_S'] = float(df_config.iloc[thresholds_start_row + 2, 1])
    results['delta'] = float(df_config.iloc[thresholds_start_row + 3, 1])
    results['theta'] = float(df_config.iloc[thresholds_start_row + 4, 1])
    results['tau_O'] = float(df_config.iloc[thresholds_start_row + 5, 1])
    results['tau_S'] = float(df_config.iloc[thresholds_start_row + 6, 1])
    results['lambda'] = float(df_config.iloc[thresholds_start_row + 7, 1])
    results['mu'] = float(df_config.iloc[thresholds_start_row + 8, 1])
    
    # Extract Step 1: Completeness
    df_comp = pd.read_excel(file, sheet_name='1_Completeness', skiprows=3, header=0)
    median_col_name = df_comp.columns[2 + num_experts]
    c_values = df_comp[median_col_name].head(num_criteria).tolist()
    results['c_values'] = c_values
    
    # Extract Step 2: Objectivity
    df_obj = pd.read_excel(file, sheet_name='2_Objectivity', skiprows=3, header=0)
    binary_col_name = df_obj.columns[4 + num_experts]
    u_values = df_obj[binary_col_name].head(num_criteria).astype(int).tolist()
    results['u_values'] = u_values
    
    # Extract Step 3: Measurability
    df_meas = pd.read_excel(file, sheet_name='3_Measurability', skiprows=3, header=0)
    median_col_name = df_meas.columns[2 + num_experts]
    m_values = df_meas[median_col_name].head(num_criteria).tolist()
    results['m_values'] = m_values
    
    # Extract Step 4: Distinctiveness
    df_dist = pd.read_excel(file, sheet_name='4_Distinctiveness', header=None)
    
    decision_matrices = []
    current_row = 4
    
    for e in range(num_experts):
        data_start = current_row + 2
        
        matrix_data = []
        for a in range(num_alternatives):
            row_data = []
            for c in range(num_criteria):
                value = df_dist.iloc[data_start + a, 1 + c]
                row_data.append(float(value) if pd.notna(value) else 0.0)
            matrix_data.append(row_data)
        
        matrix_df = pd.DataFrame(matrix_data, columns=criteria_ids)
        decision_matrices.append(matrix_df)
        
        current_row = data_start + num_alternatives + 2
    
    # Compute correlations
    correlations = []
    for matrix in decision_matrices:
        corr_matrix = matrix.corr().abs()
        correlations.append(corr_matrix.values)
    
    stacked = np.stack(correlations, axis=2)
    pooled_corr = np.median(stacked, axis=2)
    
    results['r_mat'] = pooled_corr.tolist()
    
    # Extract Step 6: Sensitivity
    df_sens = pd.read_excel(file, sheet_name='6_Sensitivity', header=None)
    
    decision_matrices_sens = []
    current_row = 4
    
    for e in range(num_experts):
        data_start = current_row + 2
        
        matrix_data = []
        for a in range(num_alternatives):
            row_data = []
            for c in range(num_criteria):
                value = df_sens.iloc[data_start + a, 1 + c]
                row_data.append(float(value) if pd.notna(value) else 0.0)
            matrix_data.append(row_data)
        
        matrix_df = pd.DataFrame(matrix_data, columns=criteria_ids)
        decision_matrices_sens.append(matrix_df)
        
        current_row = data_start + num_alternatives + 2
    
    # Normalize matrices
    def normalize_matrix(matrix, types):
        norm = matrix.copy()
        for idx, col in enumerate(matrix.columns):
            max_val = matrix[col].max()
            min_val = matrix[col].min()
            if max_val == min_val:
                norm[col] = 1.0
            elif types[idx] == 'Benefit':
                norm[col] = (matrix[col] - min_val) / (max_val - min_val)
            else:  # Cost
                norm[col] = (max_val - matrix[col]) / (max_val - min_val)
        return norm
    
    normalized_matrices = [normalize_matrix(m, criteria_types) for m in decision_matrices_sens]
    
    # Monte Carlo sensitivity
    num_simulations = 1000
    np.random.seed(42)
    random_weights = np.random.dirichlet(np.ones(num_criteria), num_simulations)
    
    sensitivity_results = []
    for norm_mat in normalized_matrices:
        elasticities = []
        for weights in random_weights:
            scores = np.dot(norm_mat.values, weights)
            total = scores.sum()
            if total > 0:
                elasticity = [(norm_mat.iloc[:, j] * weights[j]).sum() / total 
                             for j in range(num_criteria)]
            else:
                elasticity = [0.0] * num_criteria
            elasticities.append(elasticity)
        
        avg_elasticity = np.mean(elasticities, axis=0)
        sensitivity_results.append(avg_elasticity)
    
    s_values = np.mean(sensitivity_results, axis=0).tolist()
    results['s_values'] = s_values
    
    # Extract Step 7: Cost-Effectiveness
    df_cost = pd.read_excel(file, sheet_name='7_Cost_Effectiveness', skiprows=3, header=0)
    median_col_name = df_cost.columns[2 + num_experts]
    ce_values = df_cost[median_col_name].head(num_criteria).tolist()
    results['ce_values'] = ce_values
    
    # Extract Step 8: Alignment
    df_align = pd.read_excel(file, sheet_name='8_Alignment', skiprows=3, header=0)
    median_col_name = df_align.columns[2 + num_experts]
    a_values = df_align[median_col_name].head(num_criteria).tolist()
    results['a_values'] = a_values
    
    # Extract Step 9: Cognitive Coherence
    df_cog = pd.read_excel(file, sheet_name='9_Cognitive_Coherence', skiprows=3, header=0)
    num_cross_ratings = num_experts * (num_experts - 1)
    median_col_name = df_cog.columns[2 + num_cross_ratings]
    cc_values = df_cog[median_col_name].head(num_criteria).tolist()
    results['cc_values'] = cc_values
    
    # Extract Step 10: Monotone Coherence
    df_mono = pd.read_excel(file, sheet_name='10_Monotone_Coherence', skiprows=3, header=0)
    unanimity_col_name = df_mono.columns[2 + num_experts]
    q_values = df_mono[unanimity_col_name].head(num_criteria).astype(int).tolist()
    results['q_values'] = q_values
    
    # Extract Step 11: Representativeness
    df_repr = pd.read_excel(file, sheet_name='11_Representativeness', header=None)
    
    consolidated_row = None
    for idx, row in df_repr.iterrows():
        if pd.notna(row[0]) and 'CONSOLIDATED' in str(row[0]).upper():
            consolidated_row = idx + 3
            break
    
    if consolidated_row is None:
        g_matrix = np.zeros((num_criteria, num_objectives))
    else:
        g_matrix = []
        for c in range(num_criteria):
            row_data = []
            for o in range(num_objectives):
                value = df_repr.iloc[consolidated_row + c, 1 + o]
                row_data.append(int(value) if pd.notna(value) else 0)
            g_matrix.append(row_data)
        g_matrix = np.array(g_matrix)
    
    # Create objective mapping
    obj_map = {}
    for o in range(num_objectives):
        obj_idx = o + 1
        criteria_in_obj = [i + 1 for i in range(num_criteria) if g_matrix[i, o] == 1]
        if criteria_in_obj:
            obj_map[obj_idx] = criteria_in_obj
    
    results['g_matrix'] = g_matrix.tolist()
    results['obj_map'] = obj_map
    
    e_rp = (g_matrix.sum(axis=1) >= 1).astype(int).tolist()
    results['e_rp'] = e_rp
    
    Io = g_matrix.sum(axis=0).tolist()
    results['Io'] = Io
    
    # Compute derived values
    I = list(range(1, num_criteria + 1))
    O = list(range(1, num_objectives + 1))
    results['I'] = I
    results['O'] = O
    
    c = {i: results['c_values'][i-1] for i in I}
    u = {i: results['u_values'][i-1] for i in I}
    m = {i: results['m_values'][i-1] for i in I}
    s = {i: results['s_values'][i-1] for i in I}
    ce = {i: results['ce_values'][i-1] for i in I}
    a = {i: results['a_values'][i-1] for i in I}
    cc = {i: results['cc_values'][i-1] for i in I}
    q = {i: results['q_values'][i-1] for i in I}
    
    results['c'] = c
    results['u'] = u
    results['m'] = m
    results['s'] = s
    results['ce'] = ce
    results['a'] = a
    results['cc'] = cc
    results['q'] = q
    
    gamma = {i: results['gamma_O']*u[i] + results['gamma_S']*(1 - u[i]) for i in I}
    tau = {i: results['tau_O']*u[i] + results['tau_S']*(1 - u[i]) for i in I}
    results['gamma'] = gamma
    results['tau'] = tau
    
    pairs = [(i, k) for i in I for k in I if i < k]
    r_mat = results['r_mat']
    r = {(i, k): r_mat[i-1][k-1] for (i, k) in pairs}
    results['pairs'] = pairs
    results['r'] = r
    
    g = {(i, o): int(g_matrix[i-1, o-1]) for i in I for o in O}
    e_rp_dict = {i: results['e_rp'][i-1] for i in I}
    Io_dict = {o: int(results['Io'][o-1]) for o in O}
    
    results['g'] = g
    results['e_rp_dict'] = e_rp_dict
    results['Io_dict'] = Io_dict
    
    L = {o: 1 for o in O}
    U = {o: 2 for o in O}
    results['L'] = L
    results['U'] = U
    
    D = {o: max(1, Io_dict[o] - U[o]) for o in O}
    results['D'] = D
    
    results['tot_c'] = sum(c.values())
    results['tot_m'] = sum(m.values())
    results['tot_s'] = sum(s.values())
    results['tot_ce'] = sum(ce.values())
    results['tot_a'] = sum(a.values())
    results['tot_cc'] = sum(cc.values())
    results['tot_r'] = sum(r.values())
    
    results['M_big'] = 10000.0
    results['eps'] = 1e-6
    
    return results


# ================================================================
# OPTIMIZATION MODEL - COMPLETE VERSION
# ================================================================

def build_mcdm_model(data, weights):
    """Build Pyomo optimization model"""
    
    M = pyo.ConcreteModel()
    
    # Extract data
    I = data['I']
    O = data['O']
    pairs = data['pairs']
    
    c = data['c']
    u = data['u']
    m = data['m']
    s = data['s']
    ce = data['ce']
    a = data['a']
    cc = data['cc']
    q = data['q']
    
    gamma = data['gamma']
    tau = data['tau']
    r = data['r']
    g = data['g']
    e_rp = data['e_rp_dict']
    Io = data['Io_dict']
    L = data['L']
    U = data['U']
    D = data['D']
    
    alpha = data['alpha']
    delta = data['delta']
    theta = data['theta']
    lam = data['lambda']
    mu = data['mu']
    omega = data['omega']
    zeta = data['zeta']
    
    tot_c = data['tot_c']
    tot_m = data['tot_m']
    tot_s = data['tot_s']
    tot_ce = data['tot_ce']
    tot_a = data['tot_a']
    tot_cc = data['tot_cc']
    tot_r = data['tot_r']
    
    M_big = data['M_big']
    eps = data['eps']
    
    w1 = weights['w1']
    w2 = weights['w2']
    w3 = weights['w3']
    w4 = weights['w4']
    w5_minus = weights['w5_minus']
    w5_plus = weights['w5_plus']
    w6 = weights['w6']
    w7 = weights['w7']
    w8 = weights['w8']
    w9 = weights['w9']
    w11_minus = weights['w11_minus']
    w11_plus = weights['w11_plus']
    
    # Sets
    M.I = pyo.Set(initialize=I)
    M.O = pyo.Set(initialize=O)
    M.P = pyo.Set(initialize=pairs, dimen=2)
    
    # Decision variables
    M.x = pyo.Var(M.I, domain=pyo.Binary)
    
    # Property gate binaries
    M.yc = pyo.Var(M.I, domain=pyo.Binary)
    M.ym = pyo.Var(M.I, domain=pyo.Binary)
    M.ys = pyo.Var(M.I, domain=pyo.Binary)
    M.yce = pyo.Var(M.I, domain=pyo.Binary)
    M.ya = pyo.Var(M.I, domain=pyo.Binary)
    M.ycc = pyo.Var(M.I, domain=pyo.Binary)
    
    M.h = pyo.Var(M.P, domain=pyo.Binary)
    M.t = pyo.Var(M.P, domain=pyo.Binary)
    
    M.rho = pyo.Var(bounds=(0, 1))
    M.N = pyo.Var(domain=pyo.NonNegativeIntegers)
    
    M.d1_minus = pyo.Var(domain=pyo.NonNegativeIntegers)
    M.d1_plus = pyo.Var(domain=pyo.NonNegativeIntegers)
    M.d2_minus = pyo.Var(domain=pyo.NonNegativeIntegers)
    M.d2_plus = pyo.Var(domain=pyo.NonNegativeIntegers)
    
    M.n = pyo.Var(M.O, domain=pyo.NonNegativeIntegers)
    M.do1_minus = pyo.Var(M.O, domain=pyo.NonNegativeIntegers)
    M.do1_plus = pyo.Var(M.O, domain=pyo.NonNegativeIntegers)
    M.do2_minus = pyo.Var(M.O, domain=pyo.NonNegativeIntegers)
    M.do2_plus = pyo.Var(M.O, domain=pyo.NonNegativeIntegers)
    
    # Constraints
    M.comp1 = pyo.Constraint(M.I, rule=lambda M, i: c[i] - alpha <= M_big * M.yc[i] - eps)
    M.comp2 = pyo.Constraint(M.I, rule=lambda M, i: c[i] - alpha >= -M_big * (1 - M.yc[i]) - eps)
    M.comp3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.yc[i])
    
    M.N_def = pyo.Constraint(expr=M.N == sum(M.x[i] for i in M.I))
    sum_u = float(sum(u.values()))
    M.rho_def = pyo.Constraint(expr=M.rho * sum_u == sum(u[i] * M.x[i] for i in M.I))
    
    M.meas1 = pyo.Constraint(M.I, rule=lambda M, i: m[i] - gamma[i] <= M_big * M.ym[i] - eps)
    M.meas2 = pyo.Constraint(M.I, rule=lambda M, i: m[i] - gamma[i] >= -M_big * (1 - M.ym[i]) - eps)
    M.meas3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.ym[i])
    
    M.sens1 = pyo.Constraint(M.I, rule=lambda M, i: s[i] - theta <= M_big * M.ys[i])
    M.sens2 = pyo.Constraint(M.I, rule=lambda M, i: s[i] - theta >= eps - M_big * (1 - M.ys[i]))
    M.sens3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.ys[i])
    
    M.cost1 = pyo.Constraint(M.I, rule=lambda M, i: ce[i] - tau[i] <= M_big * M.yce[i] - eps)
    M.cost2 = pyo.Constraint(M.I, rule=lambda M, i: ce[i] - tau[i] >= -M_big * (1 - M.yce[i]) - eps)
    M.cost3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.yce[i])
    
    M.align1 = pyo.Constraint(M.I, rule=lambda M, i: a[i] - lam <= M_big * M.ya[i] - eps)
    M.align2 = pyo.Constraint(M.I, rule=lambda M, i: a[i] - lam >= -M_big * (1 - M.ya[i]) - eps)
    M.align3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.ya[i])
    
    M.cog1 = pyo.Constraint(M.I, rule=lambda M, i: cc[i] - mu <= M_big * M.ycc[i] - eps)
    M.cog2 = pyo.Constraint(M.I, rule=lambda M, i: cc[i] - mu >= -M_big * (1 - M.ycc[i]) - eps)
    M.cog3 = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= M.ycc[i])
    
    M.dist1 = pyo.Constraint(M.P, rule=lambda M, i, k: r[(i, k)] - delta <= M_big * M.h[(i, k)] - eps)
    M.dist2 = pyo.Constraint(M.P, rule=lambda M, i, k: r[(i, k)] - delta >= -M_big * (1 - M.h[(i, k)]) - eps)
    M.dist3 = pyo.Constraint(M.P, rule=lambda M, i, k: M.x[i] + M.x[k] <= 2 - M.h[(i, k)])
    
    M.par1 = pyo.Constraint(expr=M.N + M.d1_minus - M.d1_plus == omega)
    M.par2 = pyo.Constraint(expr=M.N + M.d2_minus - M.d2_plus == zeta)
    
    M.mono = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= q[i])
    
    M.rep_count = pyo.Constraint(M.O, rule=lambda M, o: M.n[o] == sum(g[(i, o)] * M.x[i] for i in M.I))
    M.coverage = pyo.Constraint(M.O, rule=lambda M, o: M.n[o] >= 1)
    M.rep_min = pyo.Constraint(M.O, rule=lambda M, o: M.n[o] + M.do1_minus[o] - M.do1_plus[o] == L[o])
    M.rep_max = pyo.Constraint(M.O, rule=lambda M, o: M.n[o] + M.do2_minus[o] - M.do2_plus[o] == U[o])
    M.rep_veto = pyo.Constraint(M.I, rule=lambda M, i: M.x[i] <= e_rp[i])
    
    M.t1 = pyo.Constraint(M.P, rule=lambda M, i, k: M.t[(i, k)] <= M.x[i])
    M.t2 = pyo.Constraint(M.P, rule=lambda M, i, k: M.t[(i, k)] <= M.x[k])
    M.t3 = pyo.Constraint(M.P, rule=lambda M, i, k: M.t[(i, k)] >= M.x[i] + M.x[k] - 1)
    
    # Objective
    O_card = len(O)
    
    benefit = sum(
        (
            w1 * (c[i] / tot_c) +
            w3 * (m[i] / tot_m) +
            w6 * (s[i] / tot_s) +
            w7 * (ce[i] / tot_ce) +
            w8 * (a[i] / tot_a) +
            w9 * (cc[i] / tot_cc)
        ) * M.x[i]
        for i in M.I
    )
    
    redundancy_pen = w4 * (sum(r[(i, k)] * M.t[(i, k)] for (i, k) in M.P) / tot_r)
    
    parsimony_pen = (w5_minus * (M.d1_minus / omega)) + (w5_plus * (M.d2_plus / (len(I) - zeta)))
    
    rep_pen = (
        w11_minus * (sum(M.do1_minus[o] / L[o] for o in M.O) / O_card)
        + w11_plus * (sum(M.do2_plus[o] / D[o] for o in M.O) / O_card)
    )
    
    M.obj = pyo.Objective(
        expr=benefit + w2 * M.rho - redundancy_pen - parsimony_pen - rep_pen,
        sense=pyo.maximize
    )
    
    return M


def pick_solver():
    """Select available solver"""
    for name in ("cbc", "highs", "glpk"):
        s = SolverFactory(name)
        if s.available(False):
            return s
    raise RuntimeError("No MILP solver found")


# ================================================================
# STREAMLIT UI
# ================================================================

def main():
    st.title("🎯 MCDM Criteria Selection Tool")
    st.markdown("### 11-Step Multi-Criteria Decision Analysis with Optimization")
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📊 Navigation")
        page = st.radio("", [
            "📝 1. Generate Template",
            "📤 2. Upload & Extract",
            "⚖️ 3. Set Weights",
            "🚀 4. Run Optimization"
        ])
        
        st.markdown("---")
        if st.session_state.data:
            st.markdown("**📈 Problem Info:**")
            st.metric("Criteria", st.session_state.data['num_criteria'])
            st.metric("Alternatives", st.session_state.data['num_alternatives'])
            st.metric("Experts", st.session_state.data['num_experts'])
            st.metric("Objectives", st.session_state.data['num_objectives'])
    
    # PAGE 1: GENERATE TEMPLATE
    if page == "📝 1. Generate Template":
        st.header("📝 Generate Excel Template")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Problem Structure")
            num_criteria = st.number_input("Number of Criteria", min_value=1, value=16, step=1)
            num_alternatives = st.number_input("Number of Alternatives", min_value=1, value=7, step=1)
            num_experts = st.number_input("Number of Experts", min_value=1, value=3, step=1)
            num_objectives = st.number_input("Number of Objectives", min_value=1, value=7, step=1)
        
        with col2:
            st.subheader("Parsimony Bounds")
            omega = st.number_input("Target Minimum (ω)", min_value=1, value=5, step=1)
            zeta = st.number_input("Target Maximum (ζ)", min_value=1, value=9, step=1)
        
        with st.expander("⚙️ Advanced Thresholds"):
            col1, col2 = st.columns(2)
            with col1:
                alpha = st.number_input("Completeness (α)", value=6.0)
                gamma_O = st.number_input("Measurability Objective (γ_O)", value=6.5)
                gamma_S = st.number_input("Measurability Subjective (γ_S)", value=5.5)
                delta = st.number_input("Distinctiveness (δ)", value=0.75)
                theta = st.number_input("Sensitivity (θ)", value=0.035)
            with col2:
                tau_O = st.number_input("Cost-Effectiveness Objective (τ_O)", value=7.0)
                tau_S = st.number_input("Cost-Effectiveness Subjective (τ_S)", value=6.0)
                lambda_th = st.number_input("Alignment (λ)", value=6.5)
                mu = st.number_input("Cognitive Coherence (μ)", value=7.0)
        
        if st.button("🎨 Generate Excel Template", type="primary", use_container_width=True):
            with st.spinner("Generating template..."):
                try:
                    buffer = generate_excel_template(
                        int(num_criteria), int(num_alternatives), int(num_experts), int(num_objectives),
                        int(omega), int(zeta), alpha, gamma_O, gamma_S, delta, theta,
                        tau_O, tau_S, lambda_th, mu
                    )
                    
                    st.success("✅ Template generated successfully!")
                    st.download_button(
                        label="📥 Download Excel Template",
                        data=buffer,
                        file_name=f"MCDM_Template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    # PAGE 2: UPLOAD & EXTRACT
    elif page == "📤 2. Upload & Extract":
        st.header("📤 Upload Filled Template")
        
        uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx'])
        
        if uploaded_file and st.button("🔍 Extract Data", type="primary"):
            with st.spinner("Reading Excel..."):
                try:
                    data = read_mcdm_template(uploaded_file)
                    st.session_state.data = data
                    
                    st.success("✅ Data extracted!")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Criteria", data['num_criteria'])
                    col2.metric("Alternatives", data['num_alternatives'])
                    col3.metric("Experts", data['num_experts'])
                    col4.metric("Objectives", data['num_objectives'])
                    
                    with st.expander("📋 View Criteria"):
                        for i, name in enumerate(data['criteria_names'], 1):
                            st.write(f"**C{i}:** {name} ({data['criteria_types'][i-1]})")
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    # PAGE 3: SET WEIGHTS
    elif page == "⚖️ 3. Set Weights":
        st.header("⚖️ Swing Weighting")
        
        if not st.session_state.data:
            st.warning("⚠️ Upload data first!")
            return
        
        components = {
            'w1': 'Completeness (c_i)',
            'w2': 'Objectivity (rho)',
            'w3': 'Measurability (m_i)',
            'w4': 'Distinctiveness (penalty)',
            'w5_minus': 'Parsimony lower (penalty)',
            'w6': 'Sensitivity (s_i)',
            'w7': 'Cost-Effectiveness (ce_i)',
            'w8': 'Alignment (a_i)',
            'w9': 'Cognitive Coherence (cc_i)',
            'w5_plus': 'Parsimony upper (penalty)',
            'w11_minus': 'Representativeness min (penalty)',
            'w11_plus': 'Representativeness max (penalty)'
        }
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            raw_weights = {}
            for comp_key, comp_desc in components.items():
                default = 0.1 if 'minus' not in comp_key and 'plus' not in comp_key else 0.05
                raw_weights[comp_key] = st.slider(
                    f"{comp_key}: {comp_desc}",
                    0.0, 1.0, default, 0.01
                )
        
        total = sum(raw_weights.values()) or 1
        normalized = {k: v / total for k, v in raw_weights.items()}
        st.session_state.weights = normalized
        
        with col2:
            st.subheader("Normalized Weights")
            for comp, weight in sorted(normalized.items(), key=lambda x: x[1], reverse=True):
                st.metric(comp, f"{weight:.4f}", f"{weight*100:.1f}%")
    
    # PAGE 4: RUN OPTIMIZATION
    elif page == "🚀 4. Run Optimization":
        st.header("🚀 Run Optimization")
        
        if not st.session_state.data or not st.session_state.weights:
            st.warning("⚠️ Complete previous steps first!")
            return
        
        if st.button("🎯 Run Optimization", type="primary", use_container_width=True):
            with st.spinner("Solving..."):
                try:
                    model = build_mcdm_model(st.session_state.data, st.session_state.weights)
                    solver = pick_solver()
                    result = solver.solve(model, tee=False)
                    
                    if result.solver.termination_condition == TerminationCondition.optimal:
                        st.success("✅ Optimization complete!")
                        
                        data = st.session_state.data
                        I = data['I']
                        x_val = {i: float(pyo.value(model.x[i])) for i in I}
                        selected = [i for i in I if x_val[i] > 0.5]
                        
                        st.subheader(f"✅ Selected {len(selected)} Criteria")
                        for i in selected:
                            st.write(f"**C{i}:** {data['criteria_names'][i-1]}")
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("N", int(round(pyo.value(model.N))))
                        col2.metric("ρ", f"{float(pyo.value(model.rho)):.4f}")
                        col3.metric("Objective", f"{float(pyo.value(model.obj)):.6f}")
                        
                    else:
                        st.error("❌ No solution found")
                        
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")


if __name__ == "__main__":
    main()
