import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel
import sqlite3
import pandas as pd

# Initialize the list to store selected companies
selected_companies = []

def search():
    siren = entry_siren.get()
    nom_entreprise = entry_nom_entreprise.get()

    if not siren and not nom_entreprise:
        messagebox.showwarning("Input Error", "Please enter a siren number or a company name.")
        return

    conn = sqlite3.connect('merged_database.db')
    cursor = conn.cursor()

    if siren and not nom_entreprise:
            query = """
            SELECT nom_entreprise, siren, 
                   beneficiaire_1, beneficiaire_1_date_naissance, beneficiaire_1_pourcentage_parts, beneficiaire_1_pourcentage_votes,
                   beneficiaire_2, beneficiaire_2_date_naissance, beneficiaire_2_pourcentage_parts, beneficiaire_2_pourcentage_votes,
                   beneficiaire_3, beneficiaire_3_date_naissance, beneficiaire_3_pourcentage_parts, beneficiaire_3_pourcentage_votes,
                   beneficiaire_4, beneficiaire_4_date_naissance, beneficiaire_4_pourcentage_parts, beneficiaire_4_pourcentage_votes,
                   beneficiaire_5, beneficiaire_5_date_naissance, beneficiaire_5_pourcentage_parts, beneficiaire_5_pourcentage_votes
            FROM companies 
            WHERE siren = ?
            """
            cursor.execute(query, (siren,))
    elif nom_entreprise and not siren:
            query = """
            SELECT nom_entreprise, siren, 
                   beneficiaire_1, beneficiaire_1_date_naissance, beneficiaire_1_pourcentage_parts, beneficiaire_1_pourcentage_votes,
                   beneficiaire_2, beneficiaire_2_date_naissance, beneficiaire_2_pourcentage_parts, beneficiaire_2_pourcentage_votes,
                   beneficiaire_3, beneficiaire_3_date_naissance, beneficiaire_3_pourcentage_parts, beneficiaire_3_pourcentage_votes,
                   beneficiaire_4, beneficiaire_4_date_naissance, beneficiaire_4_pourcentage_parts, beneficiaire_4_pourcentage_votes,
                   beneficiaire_5, beneficiaire_5_date_naissance, beneficiaire_5_pourcentage_parts, beneficiaire_5_pourcentage_votes
            FROM companies 
            WHERE nom_entreprise = ?
            """
            cursor.execute(query, (nom_entreprise,))
    else:
            query = """
            SELECT nom_entreprise, siren, 
                   beneficiaire_1, beneficiaire_1_date_naissance, beneficiaire_1_pourcentage_parts, beneficiaire_1_pourcentage_votes,
                   beneficiaire_2, beneficiaire_2_date_naissance, beneficiaire_2_pourcentage_parts, beneficiaire_2_pourcentage_votes,
                   beneficiaire_3, beneficiaire_3_date_naissance, beneficiaire_3_pourcentage_parts, beneficiaire_3_pourcentage_votes,
                   beneficiaire_4, beneficiaire_4_date_naissance, beneficiaire_4_pourcentage_parts, beneficiaire_4_pourcentage_votes,
                   beneficiaire_5, beneficiaire_5_date_naissance, beneficiaire_5_pourcentage_parts, beneficiaire_5_pourcentage_votes
            FROM companies 
            WHERE nom_entreprise = ? AND siren = ?
            """
            cursor.execute(query, (nom_entreprise, siren))
    results = cursor.fetchall()
    conn.close()

    if not results:
        messagebox.showinfo("No Results", "No matching entries found.")
        return

    # Open a new window to display the search results
    results_window = Toplevel(root)
    results_window.title("Search Results")

    for idx, row in enumerate(results):
        result_text = tk.Text(results_window, height=20, width=100)
        result_text.grid(row=idx, column=0, padx=10, pady=5)

        # Display the company name and siren
        result_text.insert(tk.END, f"Company: {row[0]}\nSiren: {row[1]}\n")
        
        # Display the first 5 Beneficiaries, their DOBs, and percentages
        beneficiaries_display = ""
        for i in range(2, 22, 4):  # Indices 2, 6, 10, 14, 18 for beneficiaries; 3, 7, 11, 15, 19 for DOBs; 4, 8, 12, 16, 20 for parts; 5, 9, 13, 17, 21 for votes
            if row[i]:
                beneficiaries_display += (
                    f"Beneficiary {int((i-2)/4)+1}: {row[i]} (DOB: {row[i+1]}, "
                    f"Shares: {row[i+2]}%, Votes: {row[i+3]}%)\n"
                )

        result_text.insert(tk.END, beneficiaries_display)

        add_button = tk.Button(results_window, text="Add to List",
                               command=lambda r=row: add_to_list(r))
        add_button.grid(row=idx, column=1, padx=10, pady=5)

def add_to_list(company_data):
    global selected_companies
    selected_companies.append(company_data)
    update_list_display()
    messagebox.showinfo("Added", f"Added {company_data[0]} to the list.")

def view_list():
    if not selected_companies:
        messagebox.showinfo("List Empty", "No companies added to the list.")
        return

    list_window = Toplevel(root)
    list_window.title("Compiled List")

    for idx, company in enumerate(selected_companies):
        tk.Label(list_window, text=f"{idx+1}. {company[0]} (Siren: {company[1]})").pack(padx=10, pady=5)

def clear_list():
    global selected_companies
    selected_companies.clear()  # Clear the entire list
    update_list_display()

def clear_list_item():
    global selected_companies
    if selected_companies:
        selected_companies.pop()  # Remove the last item in the list
        update_list_display()
    else:
        messagebox.showinfo("List Empty", "No companies to remove.")

def update_list_display():
    # Clear the current display in the list_frame
    for widget in list_frame.winfo_children():
        widget.destroy()

    # Add a label at the top of the list frame
    tk.Label(list_frame, text="List of Companies", bg='light green', fg='black').pack(pady=5)

    # Create numbered buttons for each selected company
    for idx, company in enumerate(selected_companies, start=1):
        company_button = tk.Button(list_frame, 
                                   text=f"{idx}. {company[0]}", 
                                   command=lambda c=company: display_company_details(c),
                                   bg='light green',  # Subtle background color
                                   fg='black',       # Subtle font color
                                   relief=tk.FLAT,   # Flat relief for a minimalist look
                                   padx=5, pady=5)
        company_button.pack(pady=5, fill=tk.X)

def display_company_details(company_data):
    # Open a new window to display the details of the selected company, similar to search results format
    details_window = Toplevel(root)
    details_window.title(company_data[0])

    result_text = tk.Text(details_window, height=20, width=100)
    result_text.grid(row=0, column=0, padx=10, pady=5)

    # Display the company name and siren
    result_text.insert(tk.END, f"Company: {company_data[0]}\nSiren: {company_data[1]}\n")
    
    # Display the first 5 Beneficiaries, their DOBs, and percentages
    beneficiaries_display = ""
    for i in range(2, 22, 4):
        if company_data[i]:
            beneficiaries_display += (
                f"Beneficiary {int((i-2)/4)+1}: {company_data[i]} (DOB: {company_data[i+1]}, "
                f"Shares: {company_data[i+2]}%, Votes: {company_data[i+3]}%)\n"
            )

    result_text.insert(tk.END, beneficiaries_display)

def import_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)

        # Check if either "Company Name" or "Siren" columns exist
        if 'Company Name' not in df.columns and 'Siren' not in df.columns:
            messagebox.showerror("Error", "Excel file must contain either a 'Company Name' or 'Siren' column.")
            return

        missing_companies = []

        # Iterate through the rows of the DataFrame
        for _, row in df.iterrows():
            siren = row.get('Siren')
            company_name = row.get('Company Name')

            conn = sqlite3.connect('merged_database.db')
            cursor = conn.cursor()

            if pd.notna(siren):  # If Siren is provided, prioritize Siren query
                cursor.execute("""
                SELECT nom_entreprise, siren, 
                       beneficiaire_1, beneficiaire_1_date_naissance, beneficiaire_1_pourcentage_parts, beneficiaire_1_pourcentage_votes,
                       beneficiaire_2, beneficiaire_2_date_naissance, beneficiaire_2_pourcentage_parts, beneficiaire_2_pourcentage_votes,
                       beneficiaire_3, beneficiaire_3_date_naissance, beneficiaire_3_pourcentage_parts, beneficiaire_3_pourcentage_votes,
                       beneficiaire_4, beneficiaire_4_date_naissance, beneficiaire_4_pourcentage_parts, beneficiaire_4_pourcentage_votes,
                       beneficiaire_5, beneficiaire_5_date_naissance, beneficiaire_5_pourcentage_parts, beneficiaire_5_pourcentage_votes
                FROM companies 
                WHERE siren = ?
                """, (siren,))
                result = cursor.fetchone()
                conn.close()

                if result:
                    selected_companies.append(result)
                else:
                    # Skip the row if Siren is provided but not found
                    continue

            elif pd.notna(company_name):  # If Siren is not provided, check by Company Name
                cursor.execute("""
                SELECT nom_entreprise, siren, 
                       beneficiaire_1, beneficiaire_1_date_naissance, beneficiaire_1_pourcentage_parts, beneficiaire_1_pourcentage_votes,
                       beneficiaire_2, beneficiaire_2_date_naissance, beneficiaire_2_pourcentage_parts, beneficiaire_2_pourcentage_votes,
                       beneficiaire_3, beneficiaire_3_date_naissance, beneficiaire_3_pourcentage_parts, beneficiaire_3_pourcentage_votes,
                       beneficiaire_4, beneficiaire_4_date_naissance, beneficiaire_4_pourcentage_parts, beneficiaire_4_pourcentage_votes,
                       beneficiaire_5, beneficiaire_5_date_naissance, beneficiaire_5_pourcentage_parts, beneficiaire_5_pourcentage_votes
                FROM companies 
                WHERE nom_entreprise = ?
                """, (company_name,))
                result = cursor.fetchone()
                conn.close()

                if result:
                    selected_companies.append(result)
                else:
                    missing_companies.append(company_name)
            else:
                continue  # Skip rows where both Company Name and Siren are missing

        update_list_display()
        if missing_companies:
            messagebox.showinfo("Import Info", f"Companies imported successfully, but the following companies were not found in the database: {', '.join(missing_companies)}")
        else:
            messagebox.showinfo("Import Success", "All companies imported successfully.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def export_list():
    if not selected_companies:
        messagebox.showinfo("List Empty", "No companies in the list to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        # Extract relevant data and column names
        data = []
        for company in selected_companies:
            company_info = {
                "Company Name": company[0],
                "Siren": company[1],
                "Beneficiary 1": company[2],
                "Beneficiary 1 DOB": company[3],
                "Beneficiary 1 Shares": company[4],
                "Beneficiary 1 Votes": company[5],
                "Beneficiary 2": company[6],
                "Beneficiary 2 DOB": company[7],
                "Beneficiary 2 Shares": company[8],
                "Beneficiary 2 Votes": company[9],
                "Beneficiary 3": company[10],
                "Beneficiary 3 DOB": company[11],
                "Beneficiary 3 Shares": company[12],
                "Beneficiary 3 Votes": company[13],
                "Beneficiary 4": company[14],
                "Beneficiary 4 DOB": company[15],
                "Beneficiary 4 Shares": company[16],
                "Beneficiary 4 Votes": company[17],
                "Beneficiary 5": company[18],
                "Beneficiary 5 DOB": company[19],
                "Beneficiary 5 Shares": company[20],
                "Beneficiary 5 Votes": company[21],
            }
            data.append(company_info)

        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Export Success", f"List exported successfully to {file_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during export: {e}")

# UI Setup
root = tk.Tk()
root.title("Company Management UI")
root.configure(bg='light green')

# Main frame that holds all widgets
main_frame = tk.Frame(root, bg='light green')
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Left frame for search and other buttons
left_frame = tk.Frame(main_frame, bg='light green')
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Right frame for the list of companies
right_frame = tk.Frame(main_frame, bg='light green')
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=20)

tk.Label(left_frame, text="Enter Siren Number:", bg='light green', fg='black').pack(pady=5)
entry_siren = tk.Entry(left_frame)
entry_siren.pack(pady=5)

tk.Label(left_frame, text="Enter Company Name:", bg='light green', fg='black').pack(pady=5)
entry_nom_entreprise = tk.Entry(left_frame)
entry_nom_entreprise.pack(pady=5)

search_button = tk.Button(left_frame, text="Search", command=search, bg='dark green', fg='white')
search_button.pack(pady=5)

view_list_button = tk.Button(left_frame, text="View List", command=view_list, bg='dark green', fg='white')
view_list_button.pack(pady=5)

clear_list_item_button = tk.Button(left_frame, text="Clear Last Entry", command=clear_list_item, bg='dark green', fg='white')
clear_list_item_button.pack(pady=5)

clear_list_button = tk.Button(left_frame, text="Clear List", command=clear_list, bg='dark green', fg='white')
clear_list_button.pack(pady=5)

import_button = tk.Button(left_frame, text="Import from Excel", command=import_excel, bg='dark green', fg='white')
import_button.pack(pady=5)

export_button = tk.Button(left_frame, text="Export List to Excel", command=export_list, bg='dark green', fg='white')
export_button.pack(pady=5)

# Frame to hold the company list as buttons
list_frame = tk.Frame(right_frame, bg='light green')
list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

root.mainloop()
