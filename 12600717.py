# CAP776: Programming in Python - Minor Project #1
# Author: Shantanu Chhetri

import openpyxl

# ============================================================
# 1. HELPER SCORING FUNCTIONS
# ============================================================
def get_feeling_score(feeling):
    feeling_scores = {
        'Excellent': 5,
        'Good': 4,
        'Neutral': 3,
        'Low': 2,
        'Stressed': 1
    }
    return feeling_scores[feeling]


def get_satisfaction_score(satisfaction):
    satisfaction_scores = {
        'Very Satisfied': 5,
        'Satisfied': 4,
        'Neutral': 3,
        'Unsatisfied': 2,
        'Very Unsatisfied': 1
    }
    return satisfaction_scores[satisfaction]


def get_energy_score(energy):
    energy_scores = {
        'High': 3,
        'Medium': 2,
        'Low': 1
    }
    return energy_scores[energy]


# ============================================================
# 2. PEARSON CORRELATION FUNCTION
# ============================================================
def calculate_correlation(x_values, y_values):
    n = len(x_values)
    if n == 0:
        return 0

    mean_x = sum(x_values) / n
    mean_y = sum(y_values) / n

    numerator = 0
    sum_x = 0
    sum_y = 0

    for i in range(n):
        x_diff = x_values[i] - mean_x
        y_diff = y_values[i] - mean_y
        numerator += x_diff * y_diff
        sum_x += x_diff ** 2
        sum_y += y_diff ** 2

    denominator = (sum_x * sum_y) ** 0.5

    if denominator == 0:
        return 0

    return numerator / denominator


# ============================================================
# 3. FILE HANDLING & DATA READING
# ============================================================
def load_activity_data(file_path, sheet_name='Daily Log'):  
    records = []
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook[sheet_name]
        
        for row in range(6, sheet.max_row + 1):
            date_val = sheet.cell(row, 1).value
            if date_val is None:
                break
                
            record = {
                'date': date_val,
                'sleep': sheet.cell(row, 2).value,
                'fitness': sheet.cell(row, 3).value,
                'study': sheet.cell(row, 4).value,
                'coding': sheet.cell(row, 5).value,
                'class': sheet.cell(row, 6).value,
                'other': sheet.cell(row, 8).value,
                'tracked': sheet.cell(row, 9).value,
                'free': sheet.cell(row, 10).value,
                'feeling': sheet.cell(row, 11).value,
                'satisfaction': sheet.cell(row, 12).value,
                'energy': sheet.cell(row, 13).value
            }
            records.append(record)
            
        print("[SUCCESS] Successfully loaded " + str(len(records)) + " records from '" + file_path + "'.")
        return records

    except FileNotFoundError:
        print("[ERROR] File '" + file_path + "' was not found. Please check the path.")
        return []
    except KeyError:
        print("[ERROR] Sheet '" + sheet_name + "' does not exist in the workbook.")
        return []
    except Exception as e:
        print("[ERROR] Failed to load data: " + str(e))
        return []


# ============================================================
# 4. DATA VALIDATION
# ============================================================
def validate_data_records(records):
    valid_records = []
    invalid_count = 0
    
    for rec in records:
        try:
            if rec['sleep'] is not None:
                valid_records.append(rec)
            else:
                invalid_count += 1
        except TypeError:
            invalid_count += 1
            
    summary = {
        'expected_days': len(records),
        'valid_days': len(valid_records),
        'missing_days': len(records) - len(valid_records),
        'invalid_records': invalid_count
    }
    return valid_records, summary


# ============================================================
# 5. CALCULATE DAILY AVERAGES
# ============================================================
def calculate_averages(valid_records):
    if not valid_records:
        print("[WARNING] No valid records available for calculation.")
        return {}

    total_sleep = 0
    total_fitness = 0
    total_study = 0
    total_coding = 0
    total_class = 0
    total_other = 0
    total_free = 0
    total_tracked = 0
    total_experience_score = 0
    
    valid_count = len(valid_records)

    for rec in valid_records:
        total_sleep += rec['sleep']
        total_fitness += rec['fitness']
        total_study += rec['study']
        total_coding += rec['coding']
        total_class += rec['class']
        total_other += rec['other']
        total_tracked += rec['tracked']
        total_free += rec['free']

        f_val = get_feeling_score(rec['feeling'])
        s_val = get_satisfaction_score(rec['satisfaction'])
        e_val = get_energy_score(rec['energy'])
        
        total_experience_score += (f_val + s_val + e_val) / 3.0

    return {
        'valid_count': valid_count,
        'avg_sleep': total_sleep / valid_count,
        'avg_fitness': total_fitness / valid_count,
        'avg_study': total_study / valid_count,
        'avg_coding': total_coding / valid_count,
        'avg_class': total_class / valid_count,
        'avg_other': total_other / valid_count,
        'avg_free': total_free / valid_count,
        'avg_tracked': total_tracked / valid_count,
        'avg_experience_score': total_experience_score / valid_count
    }


# ============================================================
# 6. CALCULATE ACTIVITY INDICES
# ============================================================
def calculate_activity_indices(averages, total_expected_days):
    if not averages:
        return {}

    TPI = averages['avg_coding']
    AAI = averages['avg_study'] + averages['avg_class']
    PhAI = averages['avg_fitness']
    SRI = averages['avg_sleep']
    ABI = averages['avg_free']
    TUI = averages['avg_tracked']
    EI = averages['avg_experience_score']
    DCI = (averages['valid_count'] / total_expected_days) * 100

    PAI = (0.15 * TPI) + (0.20 * AAI) + (0.15 * PhAI) + (0.20 * SRI) + (0.15 * TUI) + (0.10 * EI) + (0.05 * DCI)

    return {
        'TPI': TPI,
        'AAI': AAI,
        'PhAI': PhAI,
        'SRI': SRI,
        'ABI': ABI,
        'TUI': TUI,
        'EI': EI,
        'DCI': DCI,
        'PAI': PAI
    }


# ============================================================
# 7. RELATIONSHIP ANALYSIS & CORRELATIONS
# ============================================================
def analyze_relationships(valid_records):
    sleep_vals = []
    energy_vals = []
    study_vals = []
    satisfaction_vals = []
    coding_vals = []
    class_vals = []
    free_vals = []
    fitness_vals = []
    feeling_vals = []

    sleep_high_energy_days = 0
    study_high_sat_days = 0
    coding_high_energy_days = 0
    class_free_days = 0
    physical_mood_days = 0

    for rec in valid_records:
        s_val = rec['sleep']
        e_val = get_energy_score(rec['energy'])
        st_val = rec['study']
        sat_val = get_satisfaction_score(rec['satisfaction'])
        c_val = rec['coding']
        cls_val = rec['class']
        fr_val = rec['free']
        fit_val = rec['fitness']
        feel_val = get_feeling_score(rec['feeling'])

        sleep_vals.append(s_val)
        energy_vals.append(e_val)
        study_vals.append(st_val)
        satisfaction_vals.append(sat_val)
        coding_vals.append(c_val)
        class_vals.append(cls_val)
        free_vals.append(fr_val)
        fitness_vals.append(fit_val)
        feeling_vals.append(feel_val)

        if s_val >= 480 and rec['energy'] in ['High', 'Medium']:
            sleep_high_energy_days += 1
        if st_val >= 120 and rec['satisfaction'] in [ 'Very Satisfied', 'Satisfied']:
            study_high_sat_days += 1
        if c_val >= 100 and rec['energy'] in ['High', 'Medium']:
            coding_high_energy_days += 1
        if cls_val <= 250 and fr_val >= 250:
            class_free_days += 1
        if fit_val >= 30 and rec['feeling'] in ['Excellent', 'Good']:
            physical_mood_days += 1

    r_sleep_energy = calculate_correlation(sleep_vals, energy_vals)
    r_study_sat = calculate_correlation(study_vals, satisfaction_vals)
    r_coding_energy = calculate_correlation(coding_vals, energy_vals)
    r_class_free = calculate_correlation(class_vals, free_vals)
    r_fitness_feeling = calculate_correlation(fitness_vals, feeling_vals)

    return {
        'sleep_energy': str(sleep_high_energy_days) + " days with >=8 hrs sleep resulted in High/Medium energy. (r = " + str(round(r_sleep_energy, 2)) + ")",
        'study_satisfaction': str(study_high_sat_days) + " days with >=2 hrs self-study resulted in Satisfied/Highly Satisfied ratings. (r = " + str(round(r_study_sat, 2)) + ")",
        'coding_energy': str(coding_high_energy_days) + " productive coding days (>=100 mins) occurred during High/Medium energy states. (r = " + str(round(r_coding_energy, 2)) + ")",
        'class_load_free_time': str(class_free_days) + " days with moderate class time (<=250 mins) allowed >=250 mins of free time. (r = " + str(round(r_class_free, 2)) + ")",
        'physical_activity_mood': str(physical_mood_days) + " days with >=30 mins fitness time correlated with positive day feelings. (r = " + str(round(r_fitness_feeling, 2)) + ")"
    }


# ============================================================
# 8. DISPLAY REPORT FUNCTION
# ============================================================
def print_project_report(validation_summary, averages, indics, relationships):
    print("\n" + "=" * 65)
    print("   PERSONAL ACTIVITY INTELLIGENCE REPORT (CAP776 Minor Project)")
    print("=" * 65)
      
    print("\n1. ACTIVITY DATA SUMMARY")
    print("-" * 45)
    print("Expected number of days            : " + str(validation_summary['expected_days']))
    print("Valid days recorded                : " + str(validation_summary['valid_days']))
    print("Missing days                       : " + str(validation_summary['missing_days']))
    print("Invalid / excluded records         : " + str(validation_summary['invalid_records']))
    print("Average Sleep/day                  : " + str(round(averages['avg_sleep'], 2)) + " min/day")
    print("Average Fitness/day                : " + str(round(averages['avg_fitness'], 2)) + " min/day")
    print("Average Study/day                  : " + str(round(averages['avg_study'], 2)) + " min/day")
    print("Average Coding/day                 : " + str(round(averages['avg_coding'], 2)) + " min/day")
    print("Average Class/day                  : " + str(round(averages['avg_class'], 2)) + " min/day")
    print("Average Other Activities/day       : " + str(round(averages['avg_other'], 2)) + " min/day")
    print("Average Free / Unaccounted Time/day: " + str(round(averages['avg_free'], 2)) + " min/day")
      
    print("\n2. INDEX VALUES")
    print("-" * 45)
    print("Tech Productivity (TPI)            : " + str(round(indics['TPI'], 2)) + " min/day")
    print("Academic Activity (AAI)            : " + str(round(indics['AAI'], 2)) + " min/day")
    print("Physical Activity (PhAI)           : " + str(round(indics['PhAI'], 2)) + " min/day")
    print("Sleep & Recovery (SRI)             : " + str(round(indics['SRI'], 2)) + " min/day")
    print("Activity Balance (ABI)             : " + str(round(indics['ABI'], 2)) + " min/day")
    print("Time Utilization (TUI)             : " + str(round(indics['TUI'], 2)) + " min/day")
    print("Experience Index (EI)              : " + str(round(indics['EI'], 2)) + " / 5")
    print("Data Continuity Index (DCI)        : " + str(round(indics['DCI'], 2)) + " %")
    print("Personal Activity Index (PAI)      : " + str(round(indics['PAI'], 2)))
      
    print("\n3. RELATIONSHIP ANALYSIS")
    print("-" * 65)
    print("1. Sleep <-> Energy          : " + str(relationships['sleep_energy']))
    print("2. Study <-> Satisfaction    : " + str(relationships['study_satisfaction']))
    print("3. Coding <-> Energy         : " + str(relationships['coding_energy']))
    print("4. Class Time <-> Free Time  : " + str(relationships['class_load_free_time']))
    print("5. Fitness <-> Day Feeling   : " + str(relationships['physical_activity_mood']))
    print("=" * 65 + "\n")


# ============================================================
# 9. EXECUTION CONTROLLER
# ============================================================
def run_project(file_name):
    if not file_name:
        return

    raw_records = load_activity_data(file_name)
    if not raw_records:
        return

    valid_records, val_summary = validate_data_records(raw_records)
    averages = calculate_averages(valid_records)
    indices = calculate_activity_indices(averages, val_summary['expected_days'])
    relationships = analyze_relationships(valid_records)
    print_project_report(val_summary, averages, indices, relationships)


# Run program
run_project('12600717.xlsx')