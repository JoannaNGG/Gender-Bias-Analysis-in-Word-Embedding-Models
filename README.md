# Gender Bias Analysis in Word Embedding Models
This project provides tools to evaluate gender bias through analogy completion tasks and direct word bias measurements, with statistical analysis and visualisation.

**Disclaimer & Prerequisites:** 

This project was created using **Python version 3.12.10** and tested on Windows OS, it is highly recommended that you use the same as to avoid unexpected errors.

**Disk Space:** ~3-4 GB for pre-trained embeddings

**RAM:** 8BG+ recommended

The program may take 5-10 minutes to run on the first time (downloading dependencies etc.) depending on your internet connect and machine specifications, this is expected.



## Setup Instructions

1. **Clone the repository**:

```bash
git clone https://github.com/JoannaNGG/Gender-Bias-Analysis-in-Word-Embedding-Models.git
cd Gender-Bias-Analysis-in-Word-Embedding-Models
```
OR

Download as zip

2. Create and activate a Python Virtual Environment
```bash
python -m venv vevn
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```
3. Download the pre-trained embeddings

The analysis requires a large embedding file which exceeds GitHub's limits:

- Download it from Google Drive: [here](https://drive.google.com/drive/folders/1TMTN101kID6-Z77DapCDobG8EsNdyJ5l?usp=drive_link)
- Place it in the root project folder (cd Gender-Bias-Analysis-in-Word-Embedding-Models)

## Usage Guide

### 1. Analogies Evaluation ```Analogies Evaluation.py```
Analyse gender bias through analogy completion tasks

The script will:
- Prompt you to select or create an Excel output file
- Ask you to choose a text file containing analogies
- Process each analogy (A:B as C:?) through both Word2Vec and GloVe models
- Track accuracy metrics
- Calculate similarity to gendered terms (male, female, man, woman, boy, girl)
- Compute average male/female group similarities and overall bias scores
- Generate summary statistics for each model

### 2. Word List Evaluation ```Word List Evaluation.py```

Analyse gender bias for indivdual words or phrases.

The script will:
- Prompt you to select or create an Excel output file
- Ask you to choose a text file containing words/phrases to analyse
- Process each word through both models
- Calculate similarity to gendered terms (male, female, man, woman, boy, girl)
- Compute average male/female group similarities and overall bias scores
- Generate summary statistics for each model

### 3. Generate Graphs (Multiple) ```GenerateGraphs.py```

Create visualisations of bias distribution

The script will:
- Prompt you to select an Excel file
- Ask you to choose sheets for Word2Vec and GloVe results (Select the sheets named: ```Summary_modelName```)
- Generate three types of plots: Histograms, Violin plots, and Mean Point plots
- Save all graphs to ```all excel files/graphs/```
- Export statistical comparisons to  ```summary_results.xlsx```

### 4. Generate Graphs (Summary Comparison) ```GraphAll.py```

Create a combined comparison graph across multiple analyses

The script:
- Reads the ```summary_results.xlsx``` file from the graphs folder
- Creates a comparison plot showing mean bias scores with confidence intervals for analysed datasets
- Mark statistically significant differences with asterisks
- Saves the combined visualisations as ```bias_comparison``` to ```all excel files/graphs/```

## Output Format
### Analogy Evaluation Sheet Columns
- ```A, B, C```: Input words of the analogy
- ```Predicted```: Top prediction from the model
- ```Expected```: Correct answer
- ```TopScore```: Cosine similarity score of the top prediction
- ```Expected_Rank```: Rank of correct answer in top 10 results
- ```Correct_Top1/5/10```: Boolean accuracy indicators
- ```Sim_to_Male/Female/Man/Woman/Boy/Girl```: Individual similarity scores
- ```Avg_Male_Group/Female_Group```: Average similarities
- ```Gender_Bias_Score```: Computed bias score [Avg_Male_Group - Avg_Female_Group] (postive = male biased, negative = female biased)

### Word List Evaluation Sheet Columns
- ```Word```: Word/Phrase being analysed
- ```Sim_to_Male/Female/Man/Woman/Boy/Girl```
- ```Avg_Male_Group/Female_Group```
- ```Gender_Bias_Score```

### Summary Statistics
- Total number of analysed items
- Mean Bias Score
- Standard Deviation
- 95% Confidence Intervals
- Percentage of male/female leaning words
- Percentage of strongly biased words (|bias| ≥ 0.1)

## Understanding the Results
- Gender Bias Score: Calculated as **Avg_Male_Group - Avg_Female_Group**
    - Positive Scores: Word is more similar to male associated terms
    - Negative Scores: Word is more similar to female associated terms
    - Scores near zero: Relatively neutral

- Statistical Significance:
    - One sample t-test determines if mean bias differs significantly from zero
    - Mann-Whitney U test compares bias distributions between models
    - p < 0.05 indicates statistically significant bias

- Effect Size (Cohen's d):
    - ~0.2%: Small effect
    - ~0.5%: Medium effect
    - ~0.8+%: Large effect

---

**Note**: This tool is designed for research and educational purposes to understand and measure gender bias in word embeddings as part of a final year project for university. Results should be interpreted within the context of the specific models and test sets used.