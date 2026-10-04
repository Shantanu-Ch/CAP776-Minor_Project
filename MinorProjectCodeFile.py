
import openpyxl

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



def calculate_averages(valid_records):

    if not valid_records:
        print("[WARNING] No valid records available for calculation.")
        return {}

    feeling_scores = {
        'Excellent': 5,
        'Good': 4,
        'Neutral': 3,
        'Low': 2,
        'Stressed': 1
    }
    
    satisfaction_scores = {
        'Very Satisfied': 5,
        'Satisfied': 4,
        'Neutral': 3,
        'Unsatisfied': 2,
        'Very Unsatisfied': 1
    }
    
    energy_scores = {
        'High': 3,
        'Medium': 2,
        'Low': 1
    }

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

        f_val = feeling_scores[rec['feeling']]
        s_val = satisfaction_scores[rec['satisfaction']]
        e_val = energy_scores[rec['energy']]
        
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

    # Overall Personal Activity Index (PAI) calculation
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


def analyze_relationships(valid_records):
    sleep_high_energy_days = 0
    study_high_sat_days = 0
    coding_high_energy_days = 0
    class_free_days = 0
    physical_mood_days = 0
    
    positive_feelings = ['Good', 'Refreshed', 'Energetic', 'Productive', 'Relaxed']
    
    for rec in valid_records:
        if rec['sleep'] >= 480 and rec['energy'] in ['High', 'Medium']:
            sleep_high_energy_days += 1
            
        if rec['study'] >= 120 and rec['satisfaction'] in ['Satisfied', 'Highly Satisfied']:
            study_high_sat_days += 1
            
        if rec['coding'] >= 100 and rec['energy'] in ['High', 'Medium']:
            coding_high_energy_days += 1
            
        if rec['class'] <= 250 and rec['free'] >= 250:
            class_free_days += 1
            
        if rec['fitness'] >= 30 and rec['feeling'] in positive_feelings:
            physical_mood_days += 1

    return {
        'sleep_energy': str(sleep_high_energy_days) + " days with >=8 hrs sleep resulted in High/Medium energy.",
        'study_satisfaction': str(study_high_sat_days) + " days with >=2 hrs self-study resulted in Satisfied/Highly Satisfied ratings.",
        'coding_energy': str(coding_high_energy_days) + " productive coding days (>=100 mins) occurred during High/Medium energy states.",
        'class_load_free_time': str(class_free_days) + " days with moderate class time (<=250 mins) allowed >=250 mins of free time.",
        'physical_activity_mood': str(physical_mood_days) + " days with >=30 mins fitness time correlated with positive day feeling ratings."
    }


def print_project_report(validation_summary, averages, indics, relationships):

    print("\n" + "=" * 65)
    print("        PERSONAL ACTIVITY INTELLIGENCE REPORT (CAP776 MP#1)")
    print("=" * 65)
      
    print("\n1. ACTIVITY DATA SUMMARY")
    print("-" * 45)
    print(f"Expected number of days            : {validation_summary['expected_days']}")
    print(f"Valid days recorded                : {validation_summary['valid_days']}")
    print(f"Missing days                       : {validation_summary['missing_days']}")
    print(f"Invalid / excluded records         : {validation_summary['invalid_records']}")
    print(f"Average Sleep/day                  : {averages['avg_sleep']:.2f} min/day")
    print(f"Average Fitness/day                : {averages['avg_fitness']:.2f} min/day")
    print(f"Average Study/day                  : {averages['avg_study']:.2f} min/day")
    print(f"Average Coding/day                 : {averages['avg_coding']:.2f} min/day")
    print(f"Average Class/day                  : {averages['avg_class']:.2f} min/day")
    print(f"Average Other Activities/day       : {averages['avg_other']:.2f} min/day")
    print(f"Average Free / Unaccounted Time/day: {averages['avg_free']:.2f} min/day")
      
    print("\n2. INDEX VALUES")
    print("-" * 45)
    print(f"Tech Productivity (TPI)            : {indics['TPI']:.2f} min/day")
    print(f"Academic Activity (AAI)            : {indics['AAI']:.2f} min/day")
    print(f"Physical Activity (PhAI)           : {indics['PhAI']:.2f} min/day")
    print(f"Sleep & Recovery (SRI)             : {indics['SRI']:.2f} min/day")
    print(f"Activity Balance (ABI)             : {indics['ABI']:.2f} min/day")
    print(f"Time Utilization (TUI)             : {indics['TUI']:.2f} min/day")
    print(f"Experience Index (EI)              : {indics['EI']:.2f} / 5")
    print(f"Data Continuity Index (DCI)        : {indics['DCI']:.2f} %")
    print(f"Personal Activity Index (PAI)      : {indics['PAI']:.2f}")
      
    print("\n3. RELATIONSHIP ANALYSIS")
    print("-" * 45)
    print(f"1. Sleep <-> Energy       : {relationships['sleep_energy']}")
    print(f"2. Study <-> Satisfaction : {relationships['study_satisfaction']}")
    print(f"3. Coding <-> Energy      : {relationships['coding_energy']}")
    print(f"4. Class Time <-> Free Time      : {relationships['class_load_free_time']}")
    print(f"5. Fitness <-> Day Feeling      : {relationships['physical_activity_mood']}")

    print("=" * 65 + "\n")


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


run_project('12600717.xlsx')
