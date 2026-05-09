import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="MSME Loan Proposal Generator", layout="wide")

MSME_GUIDELINES = {
    "Micro": {
        "investment_limit": 10000000,
        "turnover_limit": 50000000,
        "criteria": "Investment in plant & machinery up to ₹1 Crore",
        "loan_limit": 10000000,
        "features": [
            "Lower interest rates",
            "Easy collateral requirements",
            "Quick loan processing",
            "Eligible for CGTMSE scheme",
            "Government subsidies available",
        ],
    },
    "Small": {
        "investment_limit": 100000000,
        "turnover_limit": 500000000,
        "criteria": "Investment in plant & machinery between ₹1-₹10 Crore",
        "loan_limit": 50000000,
        "features": [
            "Moderate interest rates",
            "Moderate collateral requirements",
            "Enhanced loan limits",
            "Access to technology upgradation schemes",
            "Credit guarantee coverage up to 80%",
        ],
    },
    "Medium": {
        "investment_limit": 500000000,
        "turnover_limit": 2500000000,
        "criteria": "Investment in plant & machinery between ₹10-₹50 Crore",
        "loan_limit": 100000000,
        "features": [
            "Competitive interest rates",
            "Higher loan limits",
            "Flexible repayment terms",
            "Access to capital markets",
            "Export assistance programs",
        ],
    },
}


def assess_cibil_score(score):
    if score is None:
        return None
    if score >= 750:
        return "Excellent", "A+", "Strong creditworthiness. High approval probability."
    if score >= 700:
        return "Very Good", "A", "Good credit history. Likely approval with standard terms."
    if score >= 650:
        return "Good", "B+", "Satisfactory. May require additional documentation or guarantor."
    if score >= 600:
        return "Fair", "B", "Needs improvement. Higher interest rates or lower loan amount may apply."
    if score >= 550:
        return "Poor", "C", "Weak credit profile. Significant scrutiny required. Collateral mandatory."
    return "Very Poor", "D", "Very high risk. Approval unlikely without substantial collateral or guarantor."


def rating_comment(param_name, value):
    if param_name == "profit_margin":
        if value >= 15:
            return "A", "Excellent profitability. Strong earnings potential."
        if value >= 10:
            return "B", "Good profitability. Stable income generation."
        if value >= 5:
            return "C", "Average profitability. Room for improvement."
        if value >= 0:
            return "D", "Low profitability. Risk factor for repayment capacity."
        return "E", "Negative profit. Business is unprofitable. High risk."
    if param_name == "operating_margin":
        if value >= 20:
            return "A", "Excellent operational efficiency. Costs well controlled."
        if value >= 15:
            return "B", "Good operational efficiency. Sustainable margins."
        if value >= 10:
            return "C", "Average efficiency. May face margin pressure."
        if value >= 5:
            return "D", "Low efficiency. Operational challenges present."
        return "E", "Very low/negative margins. Operational restructuring needed."
    if param_name == "roa":
        if value >= 10:
            return "A", "Excellent asset utilization. High returns on assets."
        if value >= 5:
            return "B", "Good asset utilization and returns."
        if value >= 2:
            return "C", "Average asset returns. Moderate efficiency."
        if value >= 0:
            return "D", "Low returns on assets. Asset underutilization."
        return "E", "Negative ROA. Assets not generating profits."
    if param_name == "current_ratio":
        if value >= 2.0:
            return "A", "Excellent liquidity. Strong short-term financial health."
        if value >= 1.5:
            return "B", "Good liquidity position. Can meet short-term obligations."
        if value >= 1.0:
            return "C", "Acceptable liquidity. Adequate working capital."
        if value >= 0.8:
            return "D", "Weak liquidity. Potential working capital issues."
        return "E", "Critical liquidity. Immediate solvency risk."
    if param_name == "debt_to_equity":
        if value <= 0.5:
            return "A", "Excellent leverage control. Conservative debt levels."
        if value <= 1.0:
            return "B", "Good debt-equity balance. Sustainable leverage."
        if value <= 1.5:
            return "C", "Moderate leverage. Acceptable debt levels."
        if value <= 2.0:
            return "D", "High leverage. Increased financial risk."
        return "E", "Excessive leverage. Significant solvency concerns."
    return "N/A", "Unable to assess."


def compute_financial_ratios(financials):
    profit_margin = (
        financials["net_profit"] / financials["gross_revenue"] * 100
        if financials["gross_revenue"] > 0
        else 0
    )
    operating_margin = (
        (financials["gross_revenue"] - financials["operating_expenses"]) / financials["gross_revenue"] * 100
        if financials["gross_revenue"] > 0
        else 0
    )
    roa = (
        financials["net_profit"] / financials["total_assets"] * 100
        if financials["total_assets"] > 0
        else 0
    )
    current_ratio = (
        financials["current_assets"] / financials["current_liabilities"]
        if financials["current_liabilities"] > 0
        else 0
    )
    debt_to_equity = (
        financials["total_liabilities"]
        / (financials["total_assets"] - financials["total_liabilities"])
        if (financials["total_assets"] - financials["total_liabilities"]) > 0
        else 0
    )

    return {
        "profit_margin": {
            "value": profit_margin,
            "rating": rating_comment("profit_margin", profit_margin)[0],
            "comment": rating_comment("profit_margin", profit_margin)[1],
        },
        "operating_margin": {
            "value": operating_margin,
            "rating": rating_comment("operating_margin", operating_margin)[0],
            "comment": rating_comment("operating_margin", operating_margin)[1],
        },
        "roa": {
            "value": roa,
            "rating": rating_comment("roa", roa)[0],
            "comment": rating_comment("roa", roa)[1],
        },
        "current_ratio": {
            "value": current_ratio,
            "rating": rating_comment("current_ratio", current_ratio)[0],
            "comment": rating_comment("current_ratio", current_ratio)[1],
        },
        "debt_to_equity": {
            "value": debt_to_equity,
            "rating": rating_comment("debt_to_equity", debt_to_equity)[0],
            "comment": rating_comment("debt_to_equity", debt_to_equity)[1],
        },
    }


def compute_dscr(financials, loan_requirements):
    net_operating_income = financials["gross_revenue"] - financials["operating_expenses"]
    term_loan = loan_requirements["term_loan"]
    tenure = loan_requirements["tenure_years"]
    interest_rate = loan_requirements["interest_rate"]
    annual_principal = term_loan / tenure if tenure > 0 else 0
    average_loan_balance = term_loan / 2
    annual_interest = average_loan_balance * interest_rate / 100
    total_debt_service = annual_principal + annual_interest
    dscr = net_operating_income / total_debt_service if total_debt_service > 0 else 0

    if dscr >= 1.5:
        rating = "A"
        comment = "Excellent DSCR. Very strong debt repayment capacity. Low lending risk."
    elif dscr >= 1.25:
        rating = "B"
        comment = "Good DSCR. Adequate debt repayment capacity. Acceptable lending risk."
    elif dscr >= 1.0:
        rating = "C"
        comment = "Marginal DSCR. Barely adequate to service debt. Moderate risk."
    elif dscr >= 0.8:
        rating = "D"
        comment = "Weak DSCR. Limited debt repayment capacity. High lending risk."
    else:
        rating = "E"
        comment = "Very Weak DSCR. Insufficient cash flow for debt service. Very high risk."

    return {
        "net_operating_income": net_operating_income,
        "annual_principal": annual_principal,
        "annual_interest": annual_interest,
        "total_debt_service": total_debt_service,
        "dscr": dscr,
        "rating": rating,
        "comment": comment,
    }


def assess_guarantor_profile(guarantor_details):
    assessment = {
        "has_guarantor": guarantor_details["has_guarantor"],
        "rating": "N/A",
        "comment": "No guarantor details available.",
        "support_strength": "None",
    }

    if not guarantor_details["has_guarantor"]:
        assessment["comment"] = (
            "No guarantor provided. Application will be evaluated based on the applicant's own credit, collateral, and financial profile."
        )
        return assessment

    score = guarantor_details.get("cibil_score") or 0
    net_worth = guarantor_details.get("net_worth") or 0

    if score >= 750 and net_worth >= 10000000:
        assessment["rating"] = "Strong"
        assessment["support_strength"] = "High"
        assessment["comment"] = "Excellent guarantor profile. Strong credit support and healthy net worth."
    elif score >= 700 and net_worth >= 5000000:
        assessment["rating"] = "Good"
        assessment["support_strength"] = "Moderate"
        assessment["comment"] = "Very good guarantor profile. Helpful credit support for approval."
    elif score >= 650 and net_worth >= 2000000:
        assessment["rating"] = "Fair"
        assessment["support_strength"] = "Average"
        assessment["comment"] = "Adequate guarantor support. Additional due diligence is recommended."
    elif score >= 600:
        assessment["rating"] = "Weak"
        assessment["support_strength"] = "Low"
        assessment["comment"] = "Guarantor credit is fair but may not fully mitigate risk. Verify documentation carefully."
    else:
        assessment["rating"] = "Very Weak"
        assessment["support_strength"] = "Minimal"
        assessment["comment"] = "Guarantor profile is weak. Strong collateral or improved applicant credit will be needed."

    return assessment


def determine_msme_category(annual_turnover_value):
    if annual_turnover_value < 50000000:
        return "Micro"
    if annual_turnover_value < 500000000:
        return "Small"
    return "Medium"


def generate_excel_proposal(
    business_name,
    owner_name,
    business_type,
    registration_type,
    pan_number,
    gst_number,
    annual_turnover_value,
    category,
    guarantor_details,
    guarantor_analysis,
    cibil_scores,
    financials,
    financial_ratios,
    dscr_analysis,
    loan_requirements,
    recommendation,
    bank_rating,
    bank_comments,
    provide_swot,
    swot_data,
):
    """Generate Excel file with automatic formulas"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Loan Proposal"
    
    # Define styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    subheader_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    subheader_font = Font(bold=True, size=11)
    normal_font = Font(size=10)
    currency_format = '₹#,##0.00'
    percent_format = '0.00%'
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    row = 1
    
    # Title
    ws.merge_cells(f'A{row}:D{row}')
    cell = ws[f'A{row}']
    cell.value = "COMPREHENSIVE MSME LOAN PROPOSAL"
    cell.font = Font(bold=True, size=14)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 2
    
    # Section 1: Business Information
    ws[f'A{row}'].value = "BUSINESS INFORMATION"
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].font = header_font
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    business_data = [
        ("Business Name", business_name),
        ("Owner/Proprietor Name", owner_name),
        ("Business Type", business_type),
        ("Registration Type", registration_type),
        ("PAN Number", pan_number),
        ("GST Number", gst_number),
        ("Annual Turnover (Lakhs)", annual_turnover_value / 100000),
        ("MSME Category", category),
    ]
    
    for label, value in business_data:
        ws[f'A{row}'].value = label
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = value
        if isinstance(value, float) and label == "Annual Turnover (Lakhs)":
            ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
    
    row += 1
    
    # Section 2: Guarantor Details
    ws[f'A{row}'].value = "GUARANTOR INFORMATION"
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].font = header_font
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'].value = "Has Guarantor"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = "Yes" if guarantor_details["has_guarantor"] else "No"
    row += 1
    
    if guarantor_details["has_guarantor"]:
        guarantor_data = [
            ("Guarantor Name", guarantor_details["name"]),
            ("Relationship", guarantor_details["relationship"]),
            ("Contact Number", guarantor_details["contact"]),
            ("PAN", guarantor_details["pan"]),
            ("Net Worth (Lakhs)", guarantor_details["net_worth"] / 100000),
            ("CIBIL Score", guarantor_details["cibil_score"]),
        ]
        for label, value in guarantor_data:
            ws[f'A{row}'].value = label
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'].value = value
            if isinstance(value, float) and "Lakhs" in label:
                ws[f'B{row}'].number_format = '#,##0.00'
            row += 1
    
    ws[f'A{row}'].value = "Guarantor Rating"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = guarantor_analysis["rating"]
    row += 1
    
    ws[f'A{row}'].value = "Guarantor Assessment"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = guarantor_analysis["comment"]
    ws[f'B{row}'].alignment = Alignment(wrap_text=True)
    row += 1
    
    row += 1
    
    # Section 3: Loan Requirements
    ws[f'A{row}'].value = "LOAN REQUIREMENTS"
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].font = header_font
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'].value = "Working Capital Loan (Lakhs)"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = loan_requirements["working_capital_loan"] / 100000
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 1
    
    ws[f'A{row}'].value = "Term Loan (Lakhs)"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = loan_requirements["term_loan"] / 100000
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 1
    
    ws[f'A{row}'].value = "Total Loan Required (Lakhs)"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = loan_requirements["total_loan"] / 100000
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 1
    
    ws[f'A{row}'].value = "Loan-to-Turnover Ratio (%)"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = (loan_requirements["total_loan"] / annual_turnover_value * 100) if annual_turnover_value > 0 else 0
    ws[f'B{row}'].number_format = '0.00'
    row += 1
    
    if loan_requirements["term_loan"] > 0:
        ws[f'A{row}'].value = "Term Loan Tenure (years)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = loan_requirements["tenure_years"]
        row += 1
        
        ws[f'A{row}'].value = "Interest Rate (% p.a.)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = loan_requirements["interest_rate"] / 100
        ws[f'B{row}'].number_format = '0.00%'
        row += 1
    
    row += 1
    
    # Section 4: CIBIL Scores
    if cibil_scores:
        ws[f'A{row}'].value = "CREDIT ASSESSMENT (CIBIL)"
        ws[f'A{row}'].fill = header_fill
        ws[f'A{row}'].font = header_font
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        ws[f'A{row}'].value = "Individual CIBIL Score"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = cibil_scores["individual"]
        row += 1
        
        if cibil_scores.get("commercial"):
            ws[f'A{row}'].value = "Commercial CIBIL Score"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'].value = cibil_scores["commercial"]
            row += 1
        
        row += 1
    
    # Section 5: Financial Ratios with Formulas
    if financial_ratios:
        ws[f'A{row}'].value = "FINANCIAL RATIOS & ANALYSIS"
        ws[f'A{row}'].fill = header_fill
        ws[f'A{row}'].font = header_font
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        ws[f'A{row}'].value = "Metric"
        ws[f'B{row}'].value = "Value"
        ws[f'C{row}'].value = "Rating"
        ws[f'D{row}'].value = "Comment"
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row}'].fill = subheader_fill
            ws[f'{col}{row}'].font = subheader_font
        row += 1
        
        ratio_data = [
            ("Net Profit Margin (%)", financial_ratios["profit_margin"]["value"]),
            ("Operating Margin (%)", financial_ratios["operating_margin"]["value"]),
            ("Return on Assets (%)", financial_ratios["roa"]["value"]),
            ("Current Ratio", financial_ratios["current_ratio"]["value"]),
            ("Debt-to-Equity Ratio", financial_ratios["debt_to_equity"]["value"]),
        ]
        
        for label, value in ratio_data:
            ws[f'A{row}'].value = label
            ws[f'B{row}'].value = value
            ws[f'B{row}'].number_format = '0.00'
            key = label.split()[0].lower() if "margin" in label.lower() else label.replace(" Ratio", "").replace("-", "_").lower()
            if "profit" in label.lower():
                key = "profit_margin"
            elif "operating" in label.lower():
                key = "operating_margin"
            elif "return" in label.lower():
                key = "roa"
            elif "current" in label.lower():
                key = "current_ratio"
            elif "debt" in label.lower():
                key = "debt_to_equity"
            
            ws[f'C{row}'].value = financial_ratios[key]["rating"]
            ws[f'D{row}'].value = financial_ratios[key]["comment"]
            ws[f'D{row}'].alignment = Alignment(wrap_text=True)
            row += 1
        
        row += 1
    
    # Section 6: DSCR Analysis
    if dscr_analysis:
        ws[f'A{row}'].value = "DEBT SERVICE COVERAGE RATIO (DSCR) ANALYSIS"
        ws[f'A{row}'].fill = header_fill
        ws[f'A{row}'].font = header_font
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        ws[f'A{row}'].value = "Net Operating Income (Lakhs)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = dscr_analysis["net_operating_income"] / 100000
        ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
        
        ws[f'A{row}'].value = "Annual Principal Repayment (Lakhs)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = dscr_analysis["annual_principal"] / 100000
        ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
        
        ws[f'A{row}'].value = "Annual Interest (Lakhs)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = dscr_analysis["annual_interest"] / 100000
        ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
        
        ws[f'A{row}'].value = "Total Annual Debt Service (Lakhs)"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = dscr_analysis["total_debt_service"] / 100000
        ws[f'B{row}'].number_format = '#,##0.00'
        row += 1
        
        ws[f'A{row}'].value = "DSCR"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'].value = dscr_analysis["dscr"]
        ws[f'B{row}'].number_format = '0.00'
        row += 1
        
        row += 1
    
    # Section 7: Loan Recommendation
    ws[f'A{row}'].value = "LOAN RECOMMENDATION"
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].font = header_font
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'].value = "Recommendation Score"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = f"{recommendation['score']}/100"
    row += 1
    
    ws[f'A{row}'].value = "Overall Recommendation"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = recommendation["recommendation"]
    row += 1
    
    ws[f'A{row}'].value = "Details"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = recommendation["details"]
    ws[f'B{row}'].alignment = Alignment(wrap_text=True)
    row += 1
    
    row += 1
    
    # Section 8: Bank Assessment
    ws[f'A{row}'].value = "BANK INTERNAL ASSESSMENT"
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].font = header_font
    ws.merge_cells(f'A{row}:D{row}')
    row += 1
    
    ws[f'A{row}'].value = "Bank Internal Rating"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = bank_rating
    row += 1
    
    ws[f'A{row}'].value = "Internal Comments"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].value = bank_comments
    ws[f'B{row}'].alignment = Alignment(wrap_text=True)
    row += 1
    
    # Section 9: SWOT Analysis (if provided)
    if provide_swot and swot_data:
        row += 1
        ws[f'A{row}'].value = "SWOT ANALYSIS"
        ws[f'A{row}'].fill = header_fill
        ws[f'A{row}'].font = header_font
        ws.merge_cells(f'A{row}:D{row}')
        row += 1
        
        ws[f'A{row}'].value = "Strengths"
        ws[f'A{row}'].fill = subheader_fill
        ws[f'A{row}'].font = subheader_font
        row += 1
        if "strengths" in swot_data:
            for item in swot_data["strengths"]:
                ws[f'A{row}'].value = f"• {item}"
                row += 1
        
        row += 1
        ws[f'A{row}'].value = "Weaknesses"
        ws[f'A{row}'].fill = subheader_fill
        ws[f'A{row}'].font = subheader_font
        row += 1
        if "weaknesses" in swot_data:
            for item in swot_data["weaknesses"]:
                ws[f'A{row}'].value = f"• {item}"
                row += 1
        
        row += 1
        ws[f'A{row}'].value = "Opportunities"
        ws[f'A{row}'].fill = subheader_fill
        ws[f'A{row}'].font = subheader_font
        row += 1
        if "opportunities" in swot_data:
            for item in swot_data["opportunities"]:
                ws[f'A{row}'].value = f"• {item}"
                row += 1
        
        row += 1
        ws[f'A{row}'].value = "Threats"
        ws[f'A{row}'].fill = subheader_fill
        ws[f'A{row}'].font = subheader_font
        row += 1
        if "threats" in swot_data:
            for item in swot_data["threats"]:
                ws[f'A{row}'].value = f"• {item}"
                row += 1
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 35
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def render_guidelines(category):
    guidelines = MSME_GUIDELINES[category]
    st.subheader(f"MSME Category: {category}")
    st.write(f"**Investment Limit:** ₹{guidelines['investment_limit'] / 100000:,.2f} Lakhs")
    st.write(f"**Turnover Limit:** ₹{guidelines['turnover_limit'] / 100000:,.2f} Lakhs")
    st.write(f"**Maximum Loan Amount:** ₹{guidelines['loan_limit'] / 100000:,.2f} Lakhs")
    st.write(f"**Definition:** {guidelines['criteria']}")
    st.write("**Key Features & Benefits:**")
    for feature in guidelines['features']:
        st.write(f"- {feature}")


def render_financial_analysis(financial_ratios):
    st.subheader("Financial Analysis")
    for key, value in financial_ratios.items():
        label = {
            "profit_margin": "Net Profit Margin",
            "operating_margin": "Operating Margin",
            "roa": "Return on Assets",
            "current_ratio": "Current Ratio",
            "debt_to_equity": "Debt-to-Equity Ratio",
        }[key]
        st.write(f"**{label}:** {value['value']:.2f}{'%' if key not in ['current_ratio', 'debt_to_equity'] else ''}")
        st.write(f"- Rating: {value['rating']}")
        st.write(f"- Comment: {value['comment']}")


def generate_loan_recommendation(cibil_scores, financial_ratios, dscr_analysis, loan_requirements, annual_turnover, guarantor_analysis):
    recommendation_score = 0
    factors = []

    if cibil_scores:
        ind_score = cibil_scores.get("individual")
        if ind_score is not None:
            if ind_score >= 750:
                recommendation_score += 25
                factors.append("✓ Excellent individual CIBIL score")
            elif ind_score >= 700:
                recommendation_score += 20
                factors.append("✓ Very good individual CIBIL score")
            elif ind_score >= 650:
                recommendation_score += 15
                factors.append("• Satisfactory individual CIBIL score")
            else:
                factors.append("✗ Weak individual CIBIL score - Risk factor")

    if financial_ratios:
        pm_value = financial_ratios["profit_margin"]["value"]
        cr_value = financial_ratios["current_ratio"]["value"]
        dte_value = financial_ratios["debt_to_equity"]["value"]

        if pm_value >= 15:
            recommendation_score += 15
            factors.append("✓ Excellent profit margins")
        elif pm_value >= 10:
            recommendation_score += 10
            factors.append("✓ Good profit margins")
        elif pm_value >= 5:
            recommendation_score += 5
            factors.append("• Average profit margins")
        else:
            factors.append("✗ Low profit margins - Risk factor")

        if cr_value >= 1.5:
            recommendation_score += 15
            factors.append("✓ Strong liquidity position")
        elif cr_value >= 1.0:
            recommendation_score += 10
            factors.append("• Acceptable liquidity")
        else:
            factors.append("✗ Weak liquidity - Working capital concerns")

        if dte_value <= 1.0:
            recommendation_score += 10
            factors.append("✓ Conservative leverage")
        elif dte_value <= 1.5:
            recommendation_score += 5
            factors.append("• Moderate leverage")
        else:
            factors.append("✗ High leverage - Solvency risk")

    if dscr_analysis:
        dscr = dscr_analysis["dscr"]
        if dscr >= 1.5:
            recommendation_score += 20
            factors.append("✓ Excellent debt service capacity (DSCR >= 1.5)")
        elif dscr >= 1.25:
            recommendation_score += 15
            factors.append("✓ Good debt service capacity (DSCR >= 1.25)")
        elif dscr >= 1.0:
            recommendation_score += 8
            factors.append("• Marginal debt service capacity (DSCR < 1.25)")
        else:
            factors.append("✗ Insufficient debt service capacity - High risk")

    if guarantor_analysis:
        if guarantor_analysis["has_guarantor"]:
            if guarantor_analysis["rating"] in ("Strong", "Good"):
                recommendation_score += 10
                factors.append("✓ Strong guarantor support")
            elif guarantor_analysis["rating"] == "Fair":
                recommendation_score += 5
                factors.append("• Adequate guarantor support")
            else:
                factors.append("✗ Weak guarantor support - Verify collateral and applicant credit")
        else:
            factors.append("• No guarantor provided")

    loan_to_turnover = (loan_requirements["total_loan"] / annual_turnover) * 100 if annual_turnover > 0 else 0
    if loan_to_turnover <= 25:
        recommendation_score += 10
        factors.append("✓ Conservative loan-to-turnover ratio")
    elif loan_to_turnover <= 50:
        recommendation_score += 5
        factors.append("• Moderate loan-to-turnover ratio")
    else:
        factors.append("✗ High loan-to-turnover ratio - Excessive leverage")

    if recommendation_score >= 85:
        recommendation = "APPROVED"
        details = "Strong recommendation for loan approval. Business demonstrates excellent financial health, strong repayment capacity, and good credit profile."
    elif recommendation_score >= 70:
        recommendation = "CONDITIONAL APPROVAL"
        details = "Loan can be approved subject to satisfactory guarantor, higher collateral, or minor conditions."
    elif recommendation_score >= 50:
        recommendation = "NEEDS REVIEW"
        details = "Loan requires detailed review. May be approved with additional guarantor, higher interest rate, or lower loan amount."
    else:
        recommendation = "REJECTION RECOMMENDED"
        details = "High risk profile. Not recommended for approval unless significant improvements are made or substantial guarantor support provided."

    return {
        "score": recommendation_score,
        "recommendation": recommendation,
        "details": details,
        "factors": factors,
    }


def main():
    st.title("Comprehensive MSME Loan Proposal Generator")
    st.write("A simple Streamlit deployment for MSME loan proposal preparation.")

    with st.form(key="loan_proposal_form"):
        st.header("Business & Applicant Information")
        business_name = st.text_input("Business Name")
        owner_name = st.text_input("Owner / Proprietor Name")
        business_type = st.text_input("Business Type")
        registration_type = st.text_input("Registration Type")
        pan_number = st.text_input("PAN Number")
        gst_number = st.text_input("GST Number")
        annual_turnover = st.number_input("Annual Turnover (Lakhs)", min_value=0.0, step=0.1)
        loan_purpose = st.selectbox("Purpose of Loan", ["Working Capital", "Term Loan", "Both", "Other"])
        collateral_available = st.checkbox("Collateral Available")
        collateral_details = st.text_area("Collateral Details") if collateral_available else ""

        st.header("Udyam & Promoter Profile")
        udyam_registered = st.checkbox("Udyam Registered")
        udyam_number = st.text_input("Udyam Registration Number") if udyam_registered else ""
        registration_date = st.text_input("Udyam Registration Date") if udyam_registered else ""
        age = st.text_input("Applicant Age")
        education = st.text_input("Educational Qualification")
        business_experience = st.text_input("Years of Business Experience")
        previous_loans = st.checkbox("Previous Loan Experience")
        previous_loan_details = st.text_area("Previous Loan Details") if previous_loans else ""

        st.header("Guarantor Information")
        has_guarantor = st.checkbox("Has Guarantor")
        guarantor_name = st.text_input("Guarantor Name") if has_guarantor else ""
        guarantor_relationship = st.text_input("Relationship with Applicant") if has_guarantor else ""
        guarantor_contact = st.text_input("Guarantor Contact Number") if has_guarantor else ""
        guarantor_pan = st.text_input("Guarantor PAN") if has_guarantor else ""
        guarantor_net_worth = st.number_input("Guarantor Net Worth (Lakhs)", min_value=0.0, step=0.1) if has_guarantor else 0.0
        guarantor_cibil = st.number_input("Guarantor CIBIL Score", min_value=300.0, max_value=900.0, step=1.0) if has_guarantor else 0.0

        st.header("Loan & Financial Requirements")
        working_capital_loan = st.number_input("Working Capital Loan Required (Lakhs)", min_value=0.0, step=0.1)
        term_loan = st.number_input("Term Loan Required (Lakhs)", min_value=0.0, step=0.1)
        tenure_years = st.number_input("Term Loan Tenure (years)", min_value=1, step=1)
        interest_rate = st.number_input("Expected Interest Rate (% p.a.)", min_value=0.0, step=0.1)

        st.header("Credit Information")
        provide_cibil = st.checkbox("Provide CIBIL Score Information")
        individual_cibil = st.number_input("Individual CIBIL Score", min_value=300.0, max_value=900.0, step=1.0) if provide_cibil else 0.0
        commercial_cibil = st.number_input("Commercial CIBIL Score", min_value=300.0, max_value=900.0, step=1.0) if provide_cibil else 0.0

        st.header("Financial Data")
        provide_financials = st.checkbox("Provide Detailed Financial Information")
        gross_revenue = st.number_input("Gross Annual Revenue (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        operating_expenses = st.number_input("Annual Operating Expenses (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        net_profit = st.number_input("Net Profit (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        current_assets = st.number_input("Current Assets (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        current_liabilities = st.number_input("Current Liabilities (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        total_assets = st.number_input("Total Assets (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0
        total_liabilities = st.number_input("Total Liabilities (Lakhs)", min_value=0.0, step=0.1) if provide_financials else 0.0

        st.header("SWOT Analysis")
        provide_swot = st.checkbox("Include SWOT Analysis")
        strengths = st.text_area("Strengths (comma-separated)") if provide_swot else ""
        weaknesses = st.text_area("Weaknesses (comma-separated)") if provide_swot else ""
        opportunities = st.text_area("Opportunities (comma-separated)") if provide_swot else ""
        threats = st.text_area("Threats (comma-separated)") if provide_swot else ""
        swot_data = None

        st.header("Bank Internal Assessment")
        bank_rating = st.text_input("Bank Internal Rating (e.g., A1, A2, B1, B2, C1, C2, D1, D2, E1, E2)")
        bank_comments = st.text_area("Internal Comments / Remarks")

        submitted = st.form_submit_button("Generate Proposal")

    if submitted:
        annual_turnover_value = annual_turnover * 100000
        category = determine_msme_category(annual_turnover_value)

        guarantor_details = {
            "has_guarantor": has_guarantor,
            "name": guarantor_name,
            "relationship": guarantor_relationship,
            "contact": guarantor_contact,
            "pan": guarantor_pan,
            "net_worth": guarantor_net_worth * 100000,
            "cibil_score": guarantor_cibil if has_guarantor else None,
        }
        guarantor_analysis = assess_guarantor_profile(guarantor_details)

        cibil_scores = None
        if provide_cibil:
            cibil_scores = {
                "individual": individual_cibil,
                "commercial": commercial_cibil if commercial_cibil > 0 else None,
            }

        financials = None
        financial_ratios = None
        dscr_analysis = None
        if provide_financials:
            financials = {
                "gross_revenue": gross_revenue * 100000,
                "operating_expenses": operating_expenses * 100000,
                "net_profit": net_profit * 100000,
                "current_assets": current_assets * 100000,
                "current_liabilities": current_liabilities * 100000,
                "total_assets": total_assets * 100000,
                "total_liabilities": total_liabilities * 100000,
            }
            financial_ratios = compute_financial_ratios(financials)
            if term_loan > 0:
                dscr_analysis = compute_dscr(
                    financials,
                    {
                        "term_loan": term_loan * 100000,
                        "tenure_years": tenure_years,
                        "interest_rate": interest_rate,
                    },
                )

        loan_requirements = {
            "working_capital_loan": working_capital_loan * 100000,
            "term_loan": term_loan * 100000,
            "tenure_years": tenure_years,
            "interest_rate": interest_rate,
            "total_loan": (working_capital_loan + term_loan) * 100000,
        }

        # Construct SWOT data if provided
        if provide_swot:
            swot_data = {
                "strengths": [s.strip() for s in strengths.split(',') if s.strip()],
                "weaknesses": [w.strip() for w in weaknesses.split(',') if w.strip()],
                "opportunities": [o.strip() for o in opportunities.split(',') if o.strip()],
                "threats": [t.strip() for t in threats.split(',') if t.strip()],
            }
        else:
            swot_data = None

        recommendation = generate_loan_recommendation(
            cibil_scores,
            financial_ratios,
            dscr_analysis,
            loan_requirements,
            annual_turnover_value,
            guarantor_analysis,
        )

        st.success("Proposal generated successfully")

        # Generate Excel file for download
        excel_file = generate_excel_proposal(
            business_name,
            owner_name,
            business_type,
            registration_type,
            pan_number,
            gst_number,
            annual_turnover_value,
            category,
            guarantor_details,
            guarantor_analysis,
            cibil_scores,
            financials,
            financial_ratios,
            dscr_analysis,
            loan_requirements,
            recommendation,
            bank_rating,
            bank_comments,
            provide_swot,
            swot_data if provide_swot else None,
        )
        
        # Download button
        st.download_button(
            label="📥 Download Proposal as Excel",
            data=excel_file,
            file_name=f"MSME_Loan_Proposal_{business_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("MSME Guidelines", expanded=True):
            render_guidelines(category)

        with st.expander("Applicant & Business Information", expanded=True):
            st.write(f"**Business Name:** {business_name}")
            st.write(f"**Owner Name:** {owner_name}")
            st.write(f"**Business Type:** {business_type}")
            st.write(f"**Registration Type:** {registration_type}")
            st.write(f"**PAN Number:** {pan_number}")
            st.write(f"**GST Number:** {gst_number}")
            st.write(f"**Annual Turnover:** ₹{annual_turnover_value / 100000:,.2f} Lakhs")
            st.write(f"**Loan Purpose:** {loan_purpose}")

        with st.expander("Guarantor Details", expanded=True):
            st.write(f"**Has Guarantor:** {'Yes' if has_guarantor else 'No'}")
            if has_guarantor:
                st.write(f"**Guarantor Name:** {guarantor_name}")
                st.write(f"**Relationship:** {guarantor_relationship}")
                st.write(f"**Contact:** {guarantor_contact}")
                st.write(f"**PAN:** {guarantor_pan}")
                st.write(f"**Net Worth:** ₹{guarantor_net_worth:,.2f} Lakhs")
                st.write(f"**CIBIL Score:** {guarantor_cibil}")
            st.write(f"**Guarantor Rating:** {guarantor_analysis['rating']}")
            st.write(f"**Guarantor Comments:** {guarantor_analysis['comment']}")

        with st.expander("Loan Requirement Summary", expanded=True):
            st.write(f"**Working Capital Loan:** ₹{working_capital_loan:,.2f} Lakhs")
            st.write(f"**Term Loan:** ₹{term_loan:,.2f} Lakhs")
            st.write(f"**Total Loan Required:** ₹{(working_capital_loan + term_loan):,.2f} Lakhs")
            st.write(f"**Loan-to-Turnover Ratio:** {((working_capital_loan + term_loan) / annual_turnover * 100) if annual_turnover > 0 else 0:.2f}%")
            if term_loan > 0:
                st.write(f"**Tenure:** {tenure_years} years")
                st.write(f"**Interest Rate:** {interest_rate:.2f}% p.a.")
            st.write(f"**Collateral Available:** {'Yes' if collateral_available else 'No'}")
            if collateral_available:
                st.write(f"**Collateral Details:** {collateral_details}")

        if provide_cibil:
            with st.expander("Credit Assessment (CIBIL)", expanded=True):
                ind_rating, ind_category, ind_comment = assess_cibil_score(cibil_scores['individual'])
                st.write(f"**Individual CIBIL Score:** {cibil_scores['individual']} | Category: {ind_category}")
                st.write(f"**Comment:** {ind_comment}")
                if cibil_scores.get('commercial'):
                    com_rating, com_category, com_comment = assess_cibil_score(cibil_scores['commercial'])
                    st.write(f"**Commercial CIBIL Score:** {cibil_scores['commercial']} | Category: {com_category}")
                    st.write(f"**Comment:** {com_comment}")

        if financial_ratios:
            with st.expander("Financial Ratio Analysis", expanded=True):
                render_financial_analysis(financial_ratios)

        if dscr_analysis:
            with st.expander("DSCR Analysis", expanded=True):
                st.write(f"**Net Operating Income:** ₹{dscr_analysis['net_operating_income'] / 100000:,.2f} Lakhs")
                st.write(f"**Total Annual Debt Service:** ₹{dscr_analysis['total_debt_service'] / 100000:,.2f} Lakhs")
                st.write(f"**DSCR:** {dscr_analysis['dscr']:.2f}")
                st.write(f"**Comment:** {dscr_analysis['comment']}")

        if provide_swot:
            with st.expander("SWOT Analysis", expanded=True):
                st.write("**Strengths:**")
                for item in [s.strip() for s in strengths.split(',') if s.strip()]:
                    st.write(f"- {item}")
                st.write("**Weaknesses:**")
                for item in [w.strip() for w in weaknesses.split(',') if w.strip()]:
                    st.write(f"- {item}")
                st.write("**Opportunities:**")
                for item in [o.strip() for o in opportunities.split(',') if o.strip()]:
                    st.write(f"- {item}")
                st.write("**Threats:**")
                for item in [t.strip() for t in threats.split(',') if t.strip()]:
                    st.write(f"- {item}")

        with st.expander("Loan Recommendation", expanded=True):
            st.write(f"**Recommendation Score:** {recommendation['score']}/100")
            st.write(f"**Overall Recommendation:** {recommendation['recommendation']}")
            st.write(f"**Details:** {recommendation['details']}")
            st.write("**Factors:**")
            for factor in recommendation['factors']:
                st.write(f"- {factor}")

        with st.expander("Bank Internal Assessment", expanded=True):
            st.write(f"**Bank Internal Rating:** {bank_rating}")
            st.write(f"**Internal Comments:** {bank_comments}")

        st.markdown("---")
        st.write("**Important Notes:** This proposal is a template for Indian bank loan applications. Ensure all information is accurate and verified.")
        st.write("• MSME category determines eligibility under government schemes")
        st.write("• DSCR minimum of 1.25x is typically required for term loan approval")
        st.write("• Guarantor support strengthens the application")
        st.write("• Udyam registration enhances eligibility for government benefits")


if __name__ == '__main__':
    main()