import pandas as pd
import matplotlib.pyplot as plt
import os

csat_info_path = "csat_info"
csat_years = []

def get_subject_info(subject, year):
    if year == 2017:
        if subject == "korean": return "국수영", "A:E"
        if subject == "math": return "국수영", "F:J"
    else:
        if subject == "korean": return "국수", "A:E"
        if subject == "math": return "국수", "F:J"
        if subject == "physics1": return "사과탐", "A:E"

def get_xlsx_file_paths(directory):
    files_and_folders = os.listdir(directory)
    xlsx_file_paths = [os.path.join(directory, file) for file in files_and_folders if file.endswith('.xlsx')]

    return xlsx_file_paths

def read_std_score(subject, file_path):
    year = int(os.path.basename(file_path)[:4])
    sheet, cols = get_subject_info(subject, year)

    df = pd.read_excel(file_path, sheet_name=sheet, usecols=cols)
    filtered_rows = []

    for index, row in df.iterrows():
        if row.isnull().any() or row[df.columns[0]] == "표준점수":
            continue
        if row[df.columns[0]] == '계':
            break
        filtered_rows.append(row)

    filtered_df = pd.DataFrame(filtered_rows)
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.columns = ['std_score', 'men', 'women', 'total', 'acc']

    return filtered_df

def calculate_difficulty(df):
    df['std_score'] = pd.to_numeric(df['std_score'], errors='coerce')
    df['total'] = pd.to_numeric(df['total'], errors='coerce')

    df = df.dropna(subset=['std_score', 'total'])

    total_applicants = df['acc'].iloc[-1]
    upper_limit = total_applicants * 0.04

    upper_bound = df[df['acc'] >= upper_limit]['std_score'].min()
    weighted_mean_score = (df['std_score'] * df['total']).sum() / df['total'].sum()
    std_deviation = ((df['std_score'] - weighted_mean_score) ** 2 * df['total']).sum() / df['total'].sum()
    std_deviation = std_deviation ** 0.5
    std_max = df['std_score'].iloc[0]

    return {
        'mean_score': weighted_mean_score,
        'std_deviation': std_deviation,
        'std_max': std_max,
        'upper_bound': upper_bound,
    }


def difficulty_score(test, std_max_weight=0.4, std_dev_weight=0.35, upper_bound_weight=0.3, mean_weight=0.1):
    difficulty_score = (std_max_weight * test['std_max']) + \
                       (std_dev_weight * test['std_deviation']) + \
                       (upper_bound_weight * test['upper_bound']) + \
                       (mean_weight / test['mean_score'])
    return difficulty_score


def graph_results(scores, subject):
    x_positions = range(len(csat_years))

    plt.figure(figsize=(14, 6))
    plt.plot(x_positions, scores, marker='o')
    plt.title(f'Yearly CSAT - {subject} Difficulty')
    plt.xlabel('Exams')
    plt.ylabel('Difficulty Score')
    plt.grid(True)
    plt.xticks(x_positions, csat_years)
    plt.show()

def print_results(scores, subject):
    for i, score in enumerate(scores):
        print(f'{csat_years[i]} {subject} : {int(score)}')

def find_exam_type(file_name):
    fname = list(file_name)[4:]
    if '6' in fname:
        return 1
    elif '9' in fname:
        return 2
    else:
        return 0

def main():
    # 난이도 정보 리스트
    difficulty_info_korean = []
    difficulty_info_math = []

    # 엑셀 파일 정렬 & 추출
    xlsx_files = get_xlsx_file_paths(csat_info_path)
    xlsx_files = sorted(xlsx_files)

    for file in xlsx_files:
        file_name = os.path.basename(file)
        year = file_name[:4]
        exam_type = find_exam_type(file_name)
        if exam_type==1:
            csat_years.append(year + '-6')
        elif exam_type==2:
            csat_years.append(year + '-9')
        else:
            csat_years.append(year)

    # 점수 산출
    for file in xlsx_files:
        std_korean = read_std_score("korean", file)
        difficulty_info_korean.append(calculate_difficulty(std_korean))

        std_math = read_std_score("math", file)
        difficulty_info_math.append(calculate_difficulty(std_math))

    scores_korean = [difficulty_score(test) for test in difficulty_info_korean]
    scores_math = [difficulty_score(test) for test in difficulty_info_math]

    graph_results(scores_korean, "korean")
    graph_results(scores_math, "math")


if __name__ == "__main__":
    main()
