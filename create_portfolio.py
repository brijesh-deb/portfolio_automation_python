import openpyxl
import datetime
from openpyxl.styles import PatternFill, Font

# global variables
vr_mf_start_row=0
vr_mf_end_row=0
vr_curr_row=1 #start with 1st row
vr_mf_cost=0
vr_mf_value=0
vr_mf_tot_return=0
vr_mf_cagr=0

prt_curr_row=1
prt_mf_tot_row = 0
prt_bank_tot_row= 0

file = openpyxl.load_workbook('my-investment-overviewxls-29-Aug-2021--2025.xltx')
ws = file.active
p_file = openpyxl.Workbook()
p_ws = p_file.active


# Main function which will orchestrate portfolio file creation
def create_portfolio():
    get_vr_cordinate()
    create_portfolio_header()
    create_mf_header()
    create_mf_detail()
    create_mf_summary()
    create_bank_header()
    create_bank_detail()
    create_bank_summary()
    create_networth()
    create_nw_brk_header()
    create_nw_brk_detail()
    create_nw_brk_summary()
    p_file.save("portfolio.xlsx")


def create_nw_brk_header():
    global prt_curr_row
    yellowFill = PatternFill('solid', openpyxl.styles.colors.Color(rgb="FFFF66"))
    prt_curr_row += 1 # skip 1 row
    p_ws.cell(prt_curr_row,1).value="Networth Breakup"
    p_ws.cell(prt_curr_row,2).value="Amount"
    p_ws.cell(prt_curr_row,3).value="VR%"
    # Change the format for Mutual Fund header
    for col in range(1,4):
        p_ws.cell(prt_curr_row,col).fill = yellowFill
        p_ws.cell(prt_curr_row,col).font = Font(bold=True)
    prt_curr_row += 1

def create_nw_brk_detail():
    global prt_curr_row, prt_bank_tot_row
    # Equity part of MF
    p_ws.cell(prt_curr_row,1).value="Equity - MF"
    p_ws.cell(prt_curr_row,3).value=0
    p_ws.cell(prt_curr_row,2).value='=($F${0}*$C${1})/100'.format(prt_mf_tot_row,prt_curr_row)
    prt_curr_row += 1
    # Debt part of MF
    p_ws.cell(prt_curr_row,1).value="Debt - MF"
    p_ws.cell(prt_curr_row,3).value=0
    p_ws.cell(prt_curr_row,2).value='=($F${0}*$C${1})/100'.format(prt_mf_tot_row,prt_curr_row)
    prt_curr_row += 1
    # Cash part of MF
    p_ws.cell(prt_curr_row,1).value="Cash - MF"
    p_ws.cell(prt_curr_row,3).value=0
    p_ws.cell(prt_curr_row,2).value='=($F${0}*$C${1})/100'.format(prt_mf_tot_row,prt_curr_row)
    prt_curr_row += 1
    # Other part of MF
    p_ws.cell(prt_curr_row,1).value="Other - MF"
    p_ws.cell(prt_curr_row,3).value=0
    p_ws.cell(prt_curr_row,2).value='=($F${0}*$C${1})/100'.format(prt_mf_tot_row,prt_curr_row)
    prt_curr_row += 1
    # Cash - Bank
    p_ws.cell(prt_curr_row,1).value="Cash - Bank"
    p_ws.cell(prt_curr_row,3).value=0
    p_ws.cell(prt_curr_row,2).value='=$C{0}'.format(prt_bank_tot_row)
    prt_curr_row += 1


def create_nw_brk_summary():
    pass

def create_networth():
    global prt_curr_row
    yellowFill = PatternFill('solid', openpyxl.styles.colors.Color(rgb="FFFF66"))
    prt_curr_row += 1 # skip 1 row
    p_ws.cell(prt_curr_row,1).value="Networth"
    print(prt_mf_tot_row)
    print(prt_bank_tot_row)
    p_ws.cell(prt_curr_row,2).value= '=SUM($F${0},$C${1})'.format(prt_mf_tot_row,prt_bank_tot_row)   # Formula to calculate Networth
    for col in range(1,3):
        p_ws.cell(prt_curr_row,col).fill = yellowFill
        p_ws.cell(prt_curr_row,col).font = Font(bold=True)
    prt_curr_row += 1


def create_bank_header():
    global prt_curr_row
    yellowFill = PatternFill('solid', openpyxl.styles.colors.Color(rgb="FFFF66"))
    prt_curr_row += 1 # skip 1 row
    p_ws.cell(prt_curr_row,1).value="Bank"
    p_ws.cell(prt_curr_row,2).value="A/C"
    p_ws.cell(prt_curr_row,3).value="Amt"
    # Change the format for Mutual Fund header
    for col in range(1,4):
        p_ws.cell(prt_curr_row,col).fill = yellowFill
        p_ws.cell(prt_curr_row,col).font = Font(bold=True)
    prt_curr_row += 1


def create_bank_detail():
    global prt_curr_row, prt_bank_tot_row
    p_ws.cell(prt_curr_row,1).value="ICICI"
    p_ws.cell(prt_curr_row,2).value="XXXXXXXXXX"
    p_ws.cell(prt_curr_row,3).value=0
    prt_bank_tot_row = prt_curr_row
    prt_curr_row += 1

def create_bank_summary():
    pass

def create_mf_summary():
    global prt_curr_row, prt_mf_tot_row
    prt_curr_row +=1
    p_ws.cell(prt_curr_row, 1).value = "Overall Mutual Fund"
    p_ws.cell(prt_curr_row, 5).value = ws.cell(vr_mf_end_row+1, 8).value # Total Cost
    p_ws.cell(prt_curr_row, 6).value = ws.cell(vr_mf_end_row+1, 10).value # Current Val
    p_ws.cell(prt_curr_row, 9).value = ws.cell(vr_mf_end_row+1, 14).value # CAGR
    prt_mf_tot_row = prt_curr_row
    for col in range(1,11):
        p_ws.cell(prt_curr_row,col).font = Font(bold=True)
    prt_curr_row += 1

def create_mf_detail():
    global vr_mf_start_row,vr_mf_end_row,prt_curr_row
    for row in range(vr_mf_start_row,vr_mf_end_row):
        if ws.cell(row, 11).value > 1:     # add only if Units > 1
            for col in range(1,15):
                if col==1:
                    p_ws.cell(prt_curr_row, 1).value = ws.cell(row, col).value  # Fund Name
                elif col==2:
                    p_ws.cell(prt_curr_row, 2).value = ws.cell(row, col).value  # Folio
                elif col == 11:
                    p_ws.cell(prt_curr_row, 3).value = ws.cell(row, col).value  # Units
                elif col == 9:
                    p_ws.cell(prt_curr_row, 4).value = ws.cell(row, col).value  # Cost per Units
                elif col == 8:
                    p_ws.cell(prt_curr_row, 5).value = ws.cell(row, col).value  # Total Cost
                elif col==10:
                    p_ws.cell(prt_curr_row, 6).value = ws.cell(row, col).value # Current Val
                elif col == 13:
                    p_ws.cell(prt_curr_row, 7).value = ws.cell(row, col).value  # Total Return
                elif col == 5:
                    p_ws.cell(prt_curr_row, 8).value = ws.cell(row, col).value  # NAV
                elif col == 14:
                    p_ws.cell(prt_curr_row, 9).value = ws.cell(row, col).value  # CAGR
                elif col == 3:
                    p_ws.cell(prt_curr_row, 10).value = ws.cell(row, col).value  # VR Rating
                col +=1
        row +=1
        if ws.cell(row, 11).value > 1:
            prt_curr_row +=1


# create portfolio header
def create_portfolio_header():
    global prt_curr_row
    p_ws.cell(prt_curr_row,1).value="Porfolio Owner"
    p_ws.cell(prt_curr_row,2).value="Dilip Deb"
    p_ws.cell(prt_curr_row,3).value="Period"
    p_ws.cell(prt_curr_row,4).value="Q1,FY22"
    p_ws.cell(prt_curr_row,5).value="Compiled On"
    p_ws.cell(prt_curr_row,6).value=str(datetime.date.today())
    prt_curr_row += 1

# create mutual fund header
def create_mf_header():
    global prt_curr_row
    yellowFill = PatternFill('solid', openpyxl.styles.colors.Color(rgb="FFFF66"))
    prt_curr_row += 1 # skip 1 row
    p_ws.cell(prt_curr_row,1).value="Mutual Fund"
    p_ws.cell(prt_curr_row,2).value="Folio"
    p_ws.cell(prt_curr_row,3).value="Units"
    p_ws.cell(prt_curr_row,4).value="Cost/Unit"
    p_ws.cell(prt_curr_row,5).value="Tot Cost"
    p_ws.cell(prt_curr_row,6).value="Curr Val"
    p_ws.cell(prt_curr_row,7).value="Tot Return"
    p_ws.cell(prt_curr_row,8).value="NAV"
    p_ws.cell(prt_curr_row,9).value="CAGR"
    p_ws.cell(prt_curr_row,10).value="VR"

    # Change the format for Mutual Fund header
    for col in range(1,11):
        p_ws.cell(prt_curr_row,col).fill = yellowFill
        p_ws.cell(prt_curr_row,col).font = Font(bold=True)

    prt_curr_row += 1

# Get co-ordinates of details in VR File
def get_vr_cordinate():
    # While global variables can be accessed from functions, to change their value use "global" keyword
    global vr_curr_row, vr_mf_start_row, vr_mf_end_row, vr_mf_cost,vr_mf_value,vr_mf_tot_return,vr_mf_cagr
    for row in ws.rows:
        if row[0].value == "Fund Name":
            vr_mf_start_row=vr_curr_row + 1
        elif row[0].value == "Mutual Funds":
            vr_mf_end_row = vr_curr_row-1
            vr_mf_cost=row[7].value
            vr_mf_value=row[9].value
            vr_mf_tot_return=row[12].value
            vr_mf_cagr=row[13].value
        vr_curr_row += 1


create_portfolio()

