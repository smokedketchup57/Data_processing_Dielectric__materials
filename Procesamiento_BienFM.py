import warnings
warnings.filterwarnings("ignore")

import os 
import glob 
import pandas as pd 
import numpy as np 
import io
import sys
import matplotlib.pyplot as plt

EPSILON_0 = 8.854e-14  # It is in centimeters 
pi = np.pi 

# ==============================================================================
# MULTI-BIOMASS CONFIGURATION (Dynamic Input and Output Paths)
# ==============================================================================
Biomass_Configuration = [
    {
        "Name": "Miracle Fruit",
        "Base_Path": r"./Fruta_Milagrosa",
        "Images_Path": r"./Fruta_Milagrosa/Images",
        "Excel_Name": "Consolidated_Report_Miracle_Fruit.xlsx"
    },
    {
        "Name": "Coffee Parchment",
        "Base_Path": r"./Coffee_Parchment",
        "Images_Path": r"./Coffee_Parchment/Images_CP",
        "Excel_Name": "Consolidated_Report_Coffee_Parchment.xlsx"
    }
]


# --- YOUR BASE DATA ---
base_data = [
    {"Biomass": "Miracle Fruit", "Type": "SFM", "Sample": "1st", "Diameter_cm": 1.203, "Distance_cm": 0.300},
    {"Biomass": "Miracle Fruit", "Type": "SFM", "Sample": "2nd", "Diameter_cm": 1.245, "Distance_cm": 0.285},
    {"Biomass": "Miracle Fruit", "Type": "PFM", "Sample": "1st", "Diameter_cm": 1.240, "Distance_cm": 0.250},
    {"Biomass": "Miracle Fruit", "Type": "PFM", "Sample": "2nd", "Diameter_cm": 1.235, "Distance_cm": 0.260},
    {"Biomass": "Miracle Fruit", "Type": "CFM", "Sample": "1st", "Diameter_cm": 1.255, "Distance_cm": 0.345},
    {"Biomass": "Miracle Fruit", "Type": "CFM", "Sample": "2nd", "Diameter_cm": 1.275, "Distance_cm": 0.330},
    {"Biomass": "Miracle Fruit", "Type": "SM",  "Sample": "1st", "Diameter_cm": 1.260, "Distance_cm": 0.300},
    {"Biomass": "Miracle Fruit", "Type": "SM",  "Sample": "2nd", "Diameter_cm": 1.245, "Distance_cm": 0.280},
    {"Biomass": "Miracle Fruit", "Type": "CM",  "Sample": "1st", "Diameter_cm": 1.245, "Distance_cm": 0.280},
    {"Biomass": "Miracle Fruit", "Type": "CM",  "Sample": "2nd", "Diameter_cm": 1.250, "Distance_cm": 0.300},
    {"Biomass": "Miracle Fruit", "Type": "HFM", "Sample": "1st", "Diameter_cm": 1.265, "Distance_cm": 0.315},
    {"Biomass": "Miracle Fruit", "Type": "HFM", "Sample": "2nd", "Diameter_cm": 1.265, "Distance_cm": 0.315},
    {"Biomass": "Miracle Fruit", "Type": "HM",  "Sample": "1st", "Diameter_cm": 1.250, "Distance_cm": 0.330},
    {"Biomass": "Miracle Fruit", "Type": "HM",  "Sample": "2nd", "Diameter_cm": 1.280, "Distance_cm": 0.340},
    {"Biomass": "Coffee Parchment", "Type": "CMC", "Sample": "1st", "Diameter_cm": 1.300, "Distance_cm": 0.395},
    {"Biomass": "Coffee Parchment", "Type": "CMC", "Sample": "2nd", "Diameter_cm": 1.295, "Distance_cm": 0.385},
    {"Biomass": "Coffee Parchment", "Type": "CPC", "Sample": "1st", "Diameter_cm": 1.245, "Distance_cm": 0.335},
    {"Biomass": "Coffee Parchment", "Type": "CPC", "Sample": "2nd", "Diameter_cm": 1.255, "Distance_cm": 0.325},
    {"Biomass": "Coffee Parchment", "Type": "PCT", "Sample": "1st", "Diameter_cm": 1.265, "Distance_cm": 0.345},
    {"Biomass": "Coffee Parchment", "Type": "PCT", "Sample": "2nd", "Diameter_cm": 1.265, "Distance_cm": 0.365},
]

# --- TRANSFORMATION TO QUICK SEARCH DICTIONARY ---
Pellet_Geometry = {}
for data in base_data:
    type_val = data["Type"].lower()
    sample = data["Sample"].lower()
    
    radius = data["Diameter_cm"] / 2
    area_cm2 = pi * (radius ** 2)
    
    # We add the Biomass to the dictionary to be able to use it when exporting
    Pellet_Geometry[(type_val, sample)] = {
        "Biomass": data["Biomass"],
        "A": area_cm2, 
        "D": data["Distance_cm"]
    }

print("Geometries successfully loaded into memory.\n")

# GLOBAL ACCUMULATOR: Here we will save the tables of all files

for config in Biomass_Configuration:
    print("=" * 80)
    print(f"STARTING BIOMASS ANALYSIS: {config['Name']}")
    print("=" * 80)
    
    # Environment verification: If the images folder does not exist, Python creates it
    if not os.path.exists(config["Images_Path"]):
        os.makedirs(config["Images_Path"])
        print(f"Created visual output folder: {config['Images_Path']}")
    
    # --- NEW: Verification and creation of the "Diagnostic" folder ---
    Diagnostic_Path = os.path.join(config["Base_Path"], "Diagnostic")
    if not os.path.exists(Diagnostic_Path):
        os.makedirs(Diagnostic_Path)
        print(f"Created diagnostic folder: {Diagnostic_Path}")

    # We search for the specific files of the current biomass
    MeasurementsFolder = os.path.join(config["Base_Path"], "*.xls")
    Files = glob.glob(MeasurementsFolder)
    print(f"Valid files found to process: {len(Files)}\n")

    Global_Report = []



    for i in Files: 

        FilesLower = i.lower() 
        Information = os.path.basename(FilesLower) 

        CleanName = Information.replace(".xls", "").replace("cf-", "")
        Names = CleanName.split("_")

        file_type = Names[0].strip() 
        number_str = Names[1].strip()   
        
        if "1" in number_str:
            file_sample = "1st"
        elif "2" in number_str:
            file_sample = "2nd"
        else:
            file_sample = number_str 
            
        old_stderr = sys.stderr
        old_stdout = sys.stdout
        sys.stdout = io.StringIO() 
        sys.stderr = io.StringIO()
        
        ExcelInMemory = pd.ExcelFile(i) 

        sys.stderr = old_stderr
        sys.stdout = old_stdout

        tabs = ["Data"] + [f"Append{j}" for j in range(1, 10)] 

        Pellet_Sheets_Block = []
        
        for tab in tabs: 
            try: 
                old_stderr = sys.stderr
                sys.stderr = io.StringIO()

                df_Sheet = pd.read_excel(ExcelInMemory, sheet_name=tab) 

                sys.stderr = old_stderr

                FilteredValues = df_Sheet[df_Sheet["F_AB"] >= 8000].copy() 

                Base_Columns = FilteredValues[["F_AB", "Gp_AB", "Cp_AB"]].copy() 
                Base_Columns["Gp_AB"] = Base_Columns["Gp_AB"].abs() 
                Base_Columns["Cp_AB"] = Base_Columns["Cp_AB"].abs() 

                Pellet_Sheets_Block.append(Base_Columns) 

            except Exception as e:
                sys.stderr = old_stderr
                continue


    # AVERAGING AND ERROR CALCULATION
        if Pellet_Sheets_Block:
            Temporal_Table = pd.concat(Pellet_Sheets_Block, ignore_index=True)
            
            # 1. We group by frequency and get mean and standard deviation simultaneously
            Grouped_Sample = Temporal_Table.groupby("F_AB").agg(
                MeanGp=("Gp_AB", "mean"),
                StdGp=("Gp_AB", "std"),
                MeanCp=("Cp_AB", "mean"),
                StdCp=("Cp_AB", "std")
            ).reset_index()

            Freq = Grouped_Sample["F_AB"]
            MeanGp = Grouped_Sample["MeanGp"]
            MeanCp = Grouped_Sample["MeanCp"]
            StdGp = Grouped_Sample["StdGp"]
            StdCp = Grouped_Sample["StdCp"]

            # --- DYNAMIC GEOMETRY EXTRACTION ---
            if (file_type, file_sample) in Pellet_Geometry:
                Biomass = Pellet_Geometry[(file_type, file_sample)]["Biomass"]
                A = Pellet_Geometry[(file_type, file_sample)]["A"]
                D = Pellet_Geometry[(file_type, file_sample)]["D"]
            else:
                print(f"WARNING: The pellet '{file_type}_{file_sample}' is not in the table. Skipping...")
                continue 

    # --- EXACT PHYSICAL EQUATIONS (Central Values) ---
            TangentLoss = MeanGp / (2 * pi * Freq * MeanCp)
            RealE = (MeanCp * D) / (EPSILON_0 * A)
            ImaE = (MeanGp * D) / (2 * pi * Freq * A * EPSILON_0)
            
            # NEW: AC Conductivity Calculation (S/cm)
            Conductivity = MeanGp * (D / A)
            # Uncomment the following line if you strictly need S/mm
            # Conductivity = Conductivity * 0.1 

            # --- STANDARD DEVIATION AND STANDARD ERROR (SEM) CALCULATION ---
            N = 10 # Your 10 replicates (Data + 9 Appends)
            Root_N = np.sqrt(N)

            # Errors for Real Epsilon
            Std_RealE = (StdCp * D) / (EPSILON_0 * A)
            Error_RealE = Std_RealE / Root_N

            # Errors for Imaginary Epsilon
            Std_ImaE = (StdGp * D) / (2 * pi * Freq * A * EPSILON_0)
            Error_ImaE = Std_ImaE / Root_N

            # NEW: Errors for Conductivity (Linear dependence of Gp)
            Std_Conductivity = StdGp * (D / A)
            # If you activated S/mm above, you should do: Std_Conductivity = Std_Conductivity * 0.1
            Error_Conductivity = Std_Conductivity / Root_N

            # Errors for Loss Tangent
            Relative_Error_Gp = StdGp / MeanGp
            Relative_Error_Cp = StdCp / MeanCp
            Std_TanDelta = TangentLoss * np.sqrt(Relative_Error_Gp**2 + Relative_Error_Cp**2)
            Error_TanDelta = Std_TanDelta / Root_N

            # Here starts the change in the code 
            # This part of the code must remain the same 

    # --- HORIZONTAL DATAFRAME ASSEMBLY ---
            suffix = f"{file_type.upper()}_{file_sample}"

            df_pellet_result = pd.DataFrame({
                f'Biomass_{suffix}': [Biomass] * len(Freq),
                f'Type_{suffix}': [file_type.upper()] * len(Freq),
                f'Sample_{suffix}': [file_sample] * len(Freq),
                f'Frequency_Hz_{suffix}': Freq.values,
                
                # --- NEW: RAW UNCONVERTED DATA FOR TROUBLESHOOTING ---
                f'Cp_Raw_{suffix}': MeanCp.values,
                f'Gp_Raw_{suffix}': MeanGp.values,
                
                # Calculated Central Values
                f'Tan_Delta_{suffix}': TangentLoss.values,
                f'Epsilon_Real_{suffix}': RealE.values,
                f'Epsilon_Imag_{suffix}': ImaE.values,
                f'Conductivity_{suffix}': Conductivity.values,
                
                # (The rest of your error and standard deviation columns remain exactly the same...)

                # Standard Deviation 
                f'Std_TanDelta_{suffix}': Std_TanDelta.values,
                f'Std_EpsilonReal_{suffix}': Std_RealE.values,
                f'Std_EpsilonImag_{suffix}': Std_ImaE.values,
                f'Std_Conductivity_{suffix}': Std_Conductivity.values, # NEW
                
                # Standard Error 
                f'Error_TanDelta_{suffix}': Error_TanDelta.values,
                f'Error_EpsilonReal_{suffix}': Error_RealE.values,
                f'Error_EpsilonImag_{suffix}': Error_ImaE.values,
                f'Error_Conductivity_{suffix}': Error_Conductivity.values # NEW
            })
            
            Global_Report.append(df_pellet_result)
            print(f"-> Data and Errors of {file_type.upper()} ({file_sample}) ready for consolidation.")


# ==============================================================================
# CONSOLIDATION AND FINAL EXPORT TO EXCEL (Outside loops)
# ==============================================================================
    if Global_Report:
        
        print(f"\nConcatenating pellets of {config['Name']} horizontally...")
        Mega_Final_Table = pd.concat(Global_Report, axis=1)
            
        # Dynamic Excel path based on active folder
        Output_Path = os.path.join(config["Base_Path"], config["Excel_Name"])
        Mega_Final_Table.to_excel(Output_Path, index=False)
        print(f"¡Consolidated Excel report saved in:\n{Output_Path}")
    else:
        print("No processed data found to export.")

    # ==============================================================================
    # DIAGNOSTIC MODULE: AUTOMATIC SAVING
    # ==============================================================================
    print(f"\nGenerating and saving instrumental diagnostic plots for {config['Name']}...")

    diagnostic_suffixes = [col.replace('Cp_Raw_', '') for col in Mega_Final_Table.columns if col.startswith('Cp_Raw_')]

    for suffix in diagnostic_suffixes:
        freq = Mega_Final_Table[f"Frequency_Hz_{suffix}"]
        raw_cp = Mega_Final_Table[f"Cp_Raw_{suffix}"]
        raw_gp = Mega_Final_Table[f"Gp_Raw_{suffix}"]
        label = suffix.replace("_", " ")
        
        # --- Plot A: Capacitance ---
        fig_cp = plt.figure(figsize=(7, 4))
        plt.loglog(freq, raw_cp, '-o', color='tab:blue', markersize=4, label=f"Raw Cp ({label})")
        plt.xlabel('Frequency [Hz]', fontsize=11)
        plt.ylabel('Raw Capacitance Cp [F]', fontsize=11)
        plt.title(f'Quality Control: Cp (Log-Log) - {label}')
        plt.grid(True, which="both", ls="--", alpha=0.5)
        plt.legend()
        plt.tight_layout()
        
        # Automatic saving and closing
        file_name_cp = os.path.join(Diagnostic_Path, f"Diagnostic_Cp_{suffix}.jpg")
        plt.savefig(file_name_cp, dpi=200, bbox_inches='tight')
        plt.close(fig_cp)
        
        # --- Plot B: Conductance ---
        fig_gp = plt.figure(figsize=(7, 4))
        plt.semilogx(freq, raw_gp, '-s', color='tab:orange', markersize=4, label=f"Raw Gp ({label})")
        plt.xlabel('Frequency [Hz]', fontsize=11)
        plt.ylabel('Raw Conductance Gp [S]', fontsize=11)
        plt.title(f'Quality Control: Gp (Semi-Log) - {label}')
        plt.grid(True, which="both", ls="--", alpha=0.5)
        plt.legend()
        plt.tight_layout()
        
        # Automatic saving and closing
        file_name_gp = os.path.join(Diagnostic_Path, f"Diagnostic_Gp_{suffix}.jpg")
        plt.savefig(file_name_gp, dpi=200, bbox_inches='tight')
        plt.close(fig_gp)

    # ==============================================================================
    # SCIENTIFIC VISUALIZATION MODULE (AUTOMATED PLOTS)
    # ==============================================================================
    print(f"\nStarting generation of scientific plots for {config['Name']}...")

    # We prepare the 3 canvases independently before drawing
    fig_eps, ax_eps = plt.subplots(figsize=(8, 6))
    fig_sig, ax_sig = plt.subplots(figsize=(8, 6))
    fig_tan, ax_tan = plt.subplots(figsize=(8, 6))

    # We filter the suffixes directly from the current table
    science_suffixes = [col.replace('Type_', '') for col in Mega_Final_Table.columns if col.startswith('Type_')]

    for suffix in science_suffixes:
        # Vector extraction
        freq = Mega_Final_Table[f"Frequency_Hz_{suffix}"]
        
        eps = Mega_Final_Table[f"Epsilon_Real_{suffix}"]
        err_eps = Mega_Final_Table[f"Error_EpsilonReal_{suffix}"]
        
        sig = Mega_Final_Table[f"Conductivity_{suffix}"]
        err_sig = Mega_Final_Table[f"Error_Conductivity_{suffix}"]
        
        tan = Mega_Final_Table[f"Tan_Delta_{suffix}"]
        err_tan = Mega_Final_Table[f"Error_TanDelta_{suffix}"]
        
        label = suffix.replace("_", " ")
        
        # We draw the curves (This repeats for each pellet)
        ax_eps.errorbar(freq, eps, yerr=err_eps, fmt='-o', markersize=4, capsize=3, alpha=0.8, label=label)
        ax_sig.errorbar(freq, sig, yerr=err_sig, fmt='-s', markersize=4, capsize=3, alpha=0.8, label=label)
        ax_tan.errorbar(freq, tan, yerr=err_tan, fmt='-^', markersize=4, capsize=3, alpha=0.8, label=label)
                
    # --- Aesthetic Format (OUTSIDE the suffix loop, executed only once per biomass) ---
    biomass = config["Name"]

    ax_eps.set_xscale('log')
    ax_eps.set_xlabel('Frequency [Hz]', fontsize=12)
    ax_eps.set_ylabel(r'$\epsilon_r$', fontsize=14)
    ax_eps.set_title(f'Relative Permittivity - {biomass}')
    ax_eps.legend(fontsize=9)
    ax_eps.grid(True, which="both", ls="--", alpha=0.4)

    ax_sig.set_xscale('log')
    ax_sig.set_xlabel('Frequency [Hz]', fontsize=12)
    ax_sig.set_ylabel(r'$\sigma_{ac}$ [S/cm]', fontsize=14)
    ax_sig.set_title(f'AC Conductivity - {biomass}')
    ax_sig.legend(fontsize=9)
    ax_sig.grid(True, which="both", ls="--", alpha=0.4)

    ax_tan.set_xscale('log')
    ax_tan.set_xlabel('Frequency [Hz]', fontsize=12)
    ax_tan.set_ylabel(r'$\tan(\delta)$', fontsize=14)
    ax_tan.set_title(f'Loss Tangent - {biomass}')
    ax_tan.legend(fontsize=9)
    ax_tan.grid(True, which="both", ls="--", alpha=0.4)

    # --- Dynamic Automatic Saving (OUTSIDE the suffix loop) ---
    file_name = biomass.replace(" ", "_")
    save_path = config["Images_Path"]

    fig_eps.savefig(os.path.join(save_path, f"Permittivity_{file_name}.jpg"), dpi=300, bbox_inches='tight')
    fig_sig.savefig(os.path.join(save_path, f"Conductivity_{file_name}.jpg"), dpi=300, bbox_inches='tight')
    fig_tan.savefig(os.path.join(save_path, f"LossTangent_{file_name}.jpg"), dpi=300, bbox_inches='tight')

    # We free memory
    plt.close(fig_eps)
    plt.close(fig_sig)
    plt.close(fig_tan)
