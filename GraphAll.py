import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

#prompt user to select an existing excel file
def choose_excel_file(directory):
    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]
    if not excel_files:
        print("No Excel files found in", directory)
        return None
    print("\nAvailable Excel files:")
    for i, f in enumerate(excel_files, 1):
        print(f"{i}. {f}")
    while True:
        choice = input("\nSelect file number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(excel_files):
            return os.path.join(directory, excel_files[int(choice) - 1])
        print("Invalid selection.")

#Generate comparison plot showing mean bias scores with CI for multiple comparisons between models
def generate_comparison(excel_file):
    plt.style.use("seaborn-v0_8-whitegrid")

    #Read excel file 
    df = pd.read_excel(excel_file)
    #Get model names
    group1_name = df['Group1'].iloc[0]
    group2_name = df['Group2'].iloc[0]
    #Get mean values
    mean1 = df['Mean1']
    mean2 = df['Mean2']
    #Calculate error bar lenghts
    err1 = [mean1 - df['CI_Lower1'], df['CI_Upper1'] - mean1]
    err2 = [mean2 - df['CI_Lower2'], df['CI_Upper2'] - mean2]
    #Create x axis labels and add * to labels where difference is significant
    new_labels = [f"{row['Graph_Label']} *" if row['Significant'] else row['Graph_Label'] for _, row in df.iterrows()]
    #Create the plot
    fig, ax = plt.subplots(figsize=(12, 7))
    x = np.arange(len(df))
    offset = 0.15

    #Plot W2v points
    ax.errorbar(x - offset, mean1, yerr=err1, fmt='o', capsize=8, 
                color='black', markersize=8, label=f"{group1_name} Mean ± 95% CI")
    #Plot GloVe points
    ax.errorbar(x + offset, mean2, yerr=err2, fmt='o', capsize=8, 
                color="#2E8B57", markersize=8, label=f"{group2_name} Mean ± 95% CI")
    ax.axhline(0, linestyle="--", linewidth=2, color="black", label="No Bias (0)")

    #Labels and titles
    ax.set_ylabel('Mean Gender Bias Score\n(+ Male-biased, − Female-biased)', fontsize=11)
    ax.set_title('Gender Bias Comparison: Word2Vec vs GloVe', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(new_labels, rotation=40, ha='right', fontsize=11)
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), frameon=False)
    ax.text(1.02, 0.85, "* Indicates significant\n   difference (p < 0.05)\n   (Mann-Whitney U Test)", 
            transform=ax.transAxes, fontsize=10, verticalalignment='top',
            bbox=dict(boxstyle="round,pad=0.5", facecolor='white', alpha=0.1, edgecolor='gray'))

    plt.tight_layout()
    #Save
    output_dir = os.path.dirname(excel_file)
    save_path = os.path.join(output_dir, 'bias_comparison.png')
    plt.savefig(save_path, dpi=300)
    print(f"Graph saved to: {save_path}")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_folder = os.path.join(base_dir, "all excel files", "graphs")
    excel_file = os.path.join(excel_folder, "summary_results.xlsx")

    if not os.path.exists(excel_file):
        print(f"File not found: {excel_file}")
    else:
        generate_comparison(excel_file)