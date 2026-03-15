import pandas as pd
import matplotlib.pyplot as plt
import os
from scipy import stats
from scipy.stats import t
from scipy.stats import mannwhitneyu

#Global colour setting
Mean_Line = "#2E8B57"
Neutral_Line = "#000000"

#prompt user to select an existing excel file or create a new one
def choose_excel(current_dir):
    excel_files = [f for f in os.listdir(current_dir) if f.endswith(".xlsx")]

    if not excel_files:
        print("No Excel files found.")
        exit()

    print("\nAvailable Excel files:")
    for i, f in enumerate(excel_files, 1):
        print(f"{i}. {f}")

    while True:
        choice = input("\nSelect file number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(excel_files):
            return os.path.join(current_dir, excel_files[int(choice)-1])
        print("Invalid selection.")
        
#Save summary statistics for Word2Vec vs GloVe comparison results to an Excel file
def save_results(stats1, stats2, mw_result, name1, name2, graph_label, graph_dir):
    results_file = os.path.join(graph_dir, "summary_results.xlsx")
    
    #Create a dataframe with all results
    df = pd.DataFrame([{
        "Graph_Label": graph_label,
        "Group1": name1,
        "Mean1": stats1['mean'],
        "CI_Lower1": stats1['ci_lower'],
        "CI_Upper1": stats1['ci_upper'],
        "Group2": name2,
        "Mean2": stats2['mean'],
        "CI_Lower2": stats2['ci_lower'],
        "CI_Upper2": stats2['ci_upper'],
        "Mann_Whitney_U": mw_result['U'],
        "Mann_Whitney_p": mw_result['p_value'],
        "Significant": mw_result['p_value'] < 0.05 if mw_result['p_value'] is not None else None
    }])

    #Append to existing file otherwise create new one
    if os.path.exists(results_file):
        existing_df = pd.read_excel(results_file)
        df = pd.concat([existing_df, df], ignore_index=True)

    df.to_excel(results_file, index=False)

#Compute statistics for gender bias scores in a dataframe
def compute_stats(df):
    df.columns = df.columns.str.strip()
    bias = df["Gender_Bias_Score"].dropna()

    #Basic stats
    mean = bias.mean()
    std = bias.std()
    n = len(bias)
    
    #Cohens d effect size 
    cohens_d = mean / std if std != 0 else 0

    #Calculate confidence interval and p valie if enough data is available
    if n > 1:
        _, p_value = stats.ttest_1samp(bias, 0) #one sample t test comparing to 0

        #95% confidence interval using t distribution
        sem = std / (n ** 0.5)  #standard error of the mean
        margin = t.ppf(0.975, df=n-1) * sem #margin of error
        ci_lower = mean - margin
        ci_upper = mean + margin
    else:
        p_value = 1
        ci_lower, ci_upper = None, None

    return {
        "mean": mean,
        "std": std,
        "cohens_d": cohens_d,
        "p_value": p_value,
        "n": n,
        "ci_lower": ci_lower,
        "ci_upper": ci_upper
    }

#Perform Mann-Whitney U test to compare bias scores between models
def compute_mannwhitney(df1, df2):
    data1 = df1["Gender_Bias_Score"].dropna()
    data2 = df2["Gender_Bias_Score"].dropna()

    if len(data1) > 0 and len(data2) > 0:
        u_stat, p_value = mannwhitneyu(data1, data2, alternative='two-sided')
    else:
        u_stat, p_value = None, None

    return {
        "U": u_stat,
        "p_value": p_value
    }

#Create side by side histograms with KDE overlays for both models
def plot_histogram(df1, df2, name1, name2, stats1, stats2, graph_dir, graph_label):

    fig, axes = plt.subplots(1, 2, figsize=(14, 6), sharey=True)

    all_data = pd.concat([
        df1["Gender_Bias_Score"],
        df2["Gender_Bias_Score"]
    ]).dropna()

    #Determine common x axis range by using 1st and 99 percentiles
    min_score = all_data.quantile(0.01)
    max_score = all_data.quantile(0.99)

    #Create a histogram for each model
    for i,(ax, df, name, stats_dict) in enumerate(zip(
        axes,
        [df1, df2],
        [name1, name2],
        [stats1, stats2]
    )):

        bias = df["Gender_Bias_Score"].dropna()

        #Histogram normalised to density
        ax.hist(
            bias,
            bins="auto",
            range=(min_score, max_score),
            edgecolor="black",
            alpha=0.8,
            density=True,
            label="Distribution"
        )

        #Kernel Density Estimation
        bias.plot(kind="kde", ax=ax, linewidth=2, label="Kernel Density Estimate")

        ax.set_xlim(min_score, max_score)

        #Reference line at 0 signify no bias
        ax.axvline(0, linestyle="--", linewidth=2,
                   color= Neutral_Line,
                   label="No Bias (0)")  # only label once

        # Mean line
        ax.axvline(stats_dict["mean"], linestyle="solid",
                   linewidth=3,
                   color=Mean_Line,
                   label="Mean Bias")

        #Dotted red Confidence interval lines
        if stats_dict["ci_lower"] is not None:
            ax.axvline(stats_dict["ci_lower"], linestyle=":",
                       linewidth=2,
                       color = "red",
                       label="95% Confidence Interval") 
            ax.axvline(stats_dict["ci_upper"], linestyle=":", linewidth=2, color = "red")

        ax.set_title(name)
        ax.set_xlabel("Gender Bias Score\n(+ Male-biased, − Female-biased)")

        significance = "Significant" if stats_dict["p_value"] < 0.05 else "Not Significant"

        #Statistical summary text box
        ax.annotate(
            f"N={stats_dict['n']}\nMean={stats_dict['mean']:.4f}\n"
            f"95% CI=[{stats_dict['ci_lower']:.4f}, {stats_dict['ci_upper']:.4f}]\n"
            f"Cohen's d={stats_dict['cohens_d']:.3f}\n"
            f"P={stats_dict['p_value']:.4f}\n{significance}",
            xy=(0.02, 0.98),
            xycoords="axes fraction",
            verticalalignment="top",
            bbox=dict(boxstyle="round", alpha=0.2)
        )

    axes[0].set_ylabel("Density")

    #Add legend to the right outside of axes
    axes[1].legend(
        loc='upper left',
        bbox_to_anchor=(1.02, 1),
        frameon=False
    )
    plt.suptitle(f"{graph_label} Word2Vec vs GloVe Gender Bias Distribution", fontsize=16)
    plt.tight_layout()

    #Save and display
    safe_label = graph_label.replace(" ", "_")
    plt.savefig(os.path.join(graph_dir, f"{safe_label}_histogram.png"), dpi=300)
    plt.show()

#Create violin plots showing distribution of bias scores for both models
def plot_violin(df1, df2, name1, name2, graph_dir, graph_label):

    plt.figure(figsize=(8, 7))

    data1 = df1["Gender_Bias_Score"].dropna()
    data2 = df2["Gender_Bias_Score"].dropna()

    #Violin plot
    plt.violinplot([data1, data2], showmeans=False, showmedians=False)

    #mean markers
    means = [data1.mean(), data2.mean()]
    plt.scatter([1, 2], means,
                color="white",
                edgecolor="black",
                s=100,
                zorder=3,
                label="Mean Bias")
    #Reference line at 0
    plt.axhline(0, linestyle="--", linewidth=2,
                label="No Bias (0)")

    plt.xticks([1, 2], [name1, name2])
    plt.ylabel("Gender Bias Score\n(+ Male-biased, − Female-biased)")
    plt.title(f"{graph_label} Word2Vec vs GloVe Distribution")

    plt.legend()
    plt.tight_layout()

    #Save and display
    safe_label = graph_label.replace(" ", "_")
    plt.savefig(os.path.join(graph_dir, f"{safe_label}_violin.png"), dpi=300)
    plt.show()

#Create point plot wiht error bars, show mean and cofidence intervals
def plot_meanpoint(stats1, stats2, name1, name2, graph_dir, graph_label):

    means = [stats1["mean"], stats2["mean"]]
    #Calculate error bar lenghts
    lower_errors = [
        stats1["mean"] - stats1["ci_lower"] if stats1["ci_lower"] is not None else 0,
        stats2["mean"] - stats2["ci_lower"] if stats2["ci_lower"] is not None else 0
    ]

    upper_errors = [
        stats1["ci_upper"] - stats1["mean"] if stats1["ci_upper"] is not None else 0,
        stats2["ci_upper"] - stats2["mean"] if stats2["ci_upper"] is not None else 0
    ]

    plt.figure(figsize=(6, 6))
    
    #Offset for cleaner look
    x_positions = [1, 1.2]
    #Create error bar plot
    plt.errorbar(
        x_positions,
        means,
        yerr=[lower_errors, upper_errors],
        fmt='o',
        capsize=8,
        color='black',
        markersize=8,
        label="Mean ± 95% Confidence Interval"
    )

    #Reference line at 0
    plt.axhline(0, linestyle="--", linewidth=2, label="No Bias (0)")
    plt.xticks(ticks=x_positions, labels=[name1, name2])
    plt.xlim(min(x_positions) - 0.1, max(x_positions) + 0.1)
    plt.ylabel("Mean Gender Bias Score")
    plt.title(f"{graph_label} Mean Gender Bias Comparison")

    #position legend outside of plot
    plt.legend(
        loc='upper left',
        bbox_to_anchor=(0, -0.025),
        ncol=1,
        frameon=False
    )
    plt.tight_layout()

    #Save and display
    safe_label = graph_label.replace(" ", "_")
    plt.savefig(os.path.join(graph_dir, f"{safe_label}_mean_point.png"), dpi=300)
    plt.show()    

#Main function to run gender bias comparison analysis
def main():
    plt.style.use("seaborn-v0_8-whitegrid")

    current_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "all excel files")
    graph_dir = os.path.join(current_dir, "graphs")
    os.makedirs(graph_dir, exist_ok=True)

    while True:
        #Select excel file
        file_path = choose_excel(current_dir)
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        print("\nAvailable sheets:")
        for i, sheet in enumerate(sheet_names, 1):
            print(f"{i}. {sheet}")

        def choose_sheet(prompt):
            while True:
                choice = input(prompt).strip()
                if choice.isdigit() and 1 <= int(choice) <= len(sheet_names):
                    return sheet_names[int(choice)-1]
                print("Invalid selection.")
        #Select sheets for both models to compare
        name1 = choose_sheet("Select sheet for Word2Vec: ")
        name2 = choose_sheet("Select sheet for GloVe: ")

        graph_label = input("\nEnter a label for this comparison: ").strip()

        #Read data from sheets
        df1 = pd.read_excel(file_path, sheet_name=name1)
        df2 = pd.read_excel(file_path, sheet_name=name2)

        #Compute statistics
        stats1 = compute_stats(df1)
        stats2 = compute_stats(df2)

        #Mann Whitney U test
        mw_result = compute_mannwhitney(df1, df2)

        #Save results
        save_results(stats1, stats2, mw_result, "Word2Vec", "GloVe", graph_label, graph_dir)

        #Generate plots
        plot_histogram(df1, df2, "Word2Vec", "GloVe", stats1, stats2, graph_dir, graph_label)
        plot_violin(df1, df2, "Word2Vec", "GloVe", graph_dir, graph_label)
        plot_meanpoint(stats1, stats2, "Word2Vec", "GloVe", graph_dir, graph_label)

        print(f"\nGraphs saved to: {graph_dir}")
        print(f"Summary results saved to: {os.path.join(graph_dir, 'summary_results.xlsx')}")

        again = input("\nGenerate graphs for another file? (y/n): ").strip().lower()
        if again != "y":
            print("Exiting.")
            break

if __name__ == "__main__":
    main()