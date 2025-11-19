import csv
import sys
from datetime import datetime
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('TkAgg')

def calculate_age(birth_date_str):
    try:
        birth_date = datetime.strptime(birth_date_str, '%d.%m.%Y')
        today = datetime.now()
        age = today.year - birth_date.year
        
        if (today.month, today.day) < (birth_date.month, birth_date.day):
            age -= 1
        
        return age
    except Exception as e:
        print(f"Помилка при обчисленні віку для дати {birth_date_str}: {e}")
        return 0

def get_age_category(age):

    if age < 18:
        return "younger_18"
    elif age <= 45:
        return "18-45"
    elif age <= 70:
        return "45-70"
    else:
        return "older_70"

def read_csv_data(csv_filename='employees.csv'):

    csv_path = Path(csv_filename)
    if not csv_path.exists():
        print(f"Помилка: Файл '{csv_filename}' не знайдено!")
        print(f"   Перевірте, чи файл знаходиться у поточній директорії.")
        return None
    
    try:
        print(f"Читання CSV файлу '{csv_filename}'...")
        employees = []
        
        with open(csv_filename, 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                age = calculate_age(row['Дата народження'])
                row['Вік'] = age
                row['Категорія'] = get_age_category(age)
                employees.append(row)
        
        print(f"   Прочитано {len(employees)} записів")
        print("Ok\n")
        return employees
        
    except Exception as e:
        print(f"Помилка при відкритті або читанні CSV файлу: {e}")
        return None


def analyze_gender(employees):

    gender_stats = {'Ч': 0, 'Ж': 0}
    
    for emp in employees:
        gender = emp['Стать']
        gender_stats[gender] = gender_stats.get(gender, 0) + 1
    
    return gender_stats


def analyze_age_categories(employees):

    age_stats = {
        'younger_18': 0,
        '18-45': 0,
        '45-70': 0,
        'older_70': 0
    }
    
    for emp in employees:
        category = emp['Категорія']
        age_stats[category] += 1
    
    return age_stats


def analyze_gender_by_age(employees):

    gender_age_stats = {
        'younger_18': {'Ч': 0, 'Ж': 0},
        '18-45': {'Ч': 0, 'Ж': 0},
        '45-70': {'Ч': 0, 'Ж': 0},
        'older_70': {'Ч': 0, 'Ж': 0}
    }
    
    for emp in employees:
        category = emp['Категорія']
        gender = emp['Стать']
        gender_age_stats[category][gender] += 1
    
    return gender_age_stats


def plot_gender_distribution(gender_stats):

    plt.figure(figsize=(8, 6))
    
    labels = ['Чоловіки', 'Жінки']
    sizes = [gender_stats['Ч'], gender_stats['Ж']]
    colors = ['#3498db', '#e74c3c']
    explode = (0.05, 0.05)
    
    plt.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=90)
    plt.title('Розподіл співробітників за статтю', fontsize=14, fontweight='bold', pad=20)
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig('gender_distribution.png', dpi=150, bbox_inches='tight')
    print("   Діаграму збережено: gender_distribution.png")


def plot_age_distribution(age_stats):

    plt.figure(figsize=(10, 6))
    
    categories = ['До 18', '18-45', '45-70', 'Старше 70']
    values = [
        age_stats['younger_18'],
        age_stats['18-45'],
        age_stats['45-70'],
        age_stats['older_70']
    ]
    colors = ['#9b59b6', '#3498db', '#2ecc71', '#f39c12']
    
    bars = plt.bar(categories, values, color=colors, edgecolor='black', linewidth=1.2)
    
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height,
                f'{int(height)}',
                ha='center', va='bottom', fontsize=11, fontweight='bold')
    
    plt.xlabel('Вікові категорії', fontsize=12, fontweight='bold')
    plt.ylabel('Кількість співробітників', fontsize=12, fontweight='bold')
    plt.title('Розподіл співробітників за віковими категоріями', fontsize=14, fontweight='bold', pad=20)
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    plt.savefig('age_distribution.png', dpi=150, bbox_inches='tight')
    print("   Діаграму збережено: age_distribution.png")


def plot_gender_by_age_categories(gender_age_stats):

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle('Розподіл за статтю в кожній віковій категорії', 
                 fontsize=16, fontweight='bold', y=0.995)
    
    categories = [
        ('younger_18', 'До 18 років', axes[0, 0]),
        ('18-45', '18-45 років', axes[0, 1]),
        ('45-70', '45-70 років', axes[1, 0]),
        ('older_70', 'Старше 70 років', axes[1, 1])
    ]
    
    for cat_key, cat_name, ax in categories:
        data = gender_age_stats[cat_key]
        labels = ['Чоловіки', 'Жінки']
        sizes = [data['Ч'], data['Ж']]
        colors = ['#3498db', '#e74c3c']
        
        if sum(sizes) > 0:
            ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                   shadow=True, startangle=90)
            ax.set_title(f'{cat_name}\n(Всього: {sum(sizes)})', 
                        fontsize=12, fontweight='bold', pad=10)
        else:
            ax.text(0.5, 0.5, 'Немає даних', ha='center', va='center',
                   fontsize=12, transform=ax.transAxes)
            ax.set_title(cat_name, fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('gender_by_age_distribution.png', dpi=150, bbox_inches='tight')
    print("   Діаграму збережено: gender_by_age_distribution.png")


def print_statistics(gender_stats, age_stats, gender_age_stats):

    total = gender_stats['Ч'] + gender_stats['Ж']
    
    print("=" * 70)
    print("1. СТАТИСТИКА ЗА СТАТТЮ")
    print("=" * 70)
    print(f"   - Чоловіків: {gender_stats['Ч']} ({gender_stats['Ч']/total*100:.1f}%)")
    print(f"   - Жінок: {gender_stats['Ж']} ({gender_stats['Ж']/total*100:.1f}%)")
    print(f"   - Всього: {total}")
    print()
    
    print("=" * 70)
    print("2. СТАТИСТИКА ЗА ВІКОВИМИ КАТЕГОРІЯМИ")
    print("=" * 70)
    print(f"   - До 18 років: {age_stats['younger_18']} ({age_stats['younger_18']/total*100:.1f}%)")
    print(f"   - 18-45 років: {age_stats['18-45']} ({age_stats['18-45']/total*100:.1f}%)")
    print(f"   - 45-70 років: {age_stats['45-70']} ({age_stats['45-70']/total*100:.1f}%)")
    print(f"   - Старше 70 років: {age_stats['older_70']} ({age_stats['older_70']/total*100:.1f}%)")
    print(f"   - Всього: {total}")
    print()
    
    print("=" * 70)
    print("3. СТАТИСТИКА ЗА СТАТТЮ В КОЖНІЙ ВІКОВІЙ КАТЕГОРІЇ")
    print("=" * 70)
    
    age_cat_names = {
        'younger_18': 'До 18 років',
        '18-45': '18-45 років',
        '45-70': '45-70 років',
        'older_70': 'Старше 70 років'
    }
    
    for cat_key in ['younger_18', '18-45', '45-70', 'older_70']:
        cat_name = age_cat_names[cat_key]
        data = gender_age_stats[cat_key]
        cat_total = data['Ч'] + data['Ж']
        
        print(f"\n   {cat_name}:")
        if cat_total > 0:
            print(f"      - Чоловіків: {data['Ч']} ({data['Ч']/cat_total*100:.1f}%)")
            print(f"      - Жінок: {data['Ж']} ({data['Ж']/cat_total*100:.1f}%)")
            print(f"      - Всього: {cat_total}")
        else:
            print(f"      - Немає даних")
    print()


def main():
    print("=" * 70)
    print("Програма 3 - Аналіз даних співробітників та побудова діаграм")
    print("=" * 70)
    print()
    
    # Читання CSV файлу
    employees = read_csv_data('employees.csv')
    
    if employees is None:
        sys.exit(1)
    
    # Аналіз даних
    print("Аналіз даних...")
    gender_stats = analyze_gender(employees)
    age_stats = analyze_age_categories(employees)
    gender_age_stats = analyze_gender_by_age(employees)
    print()
    
    # Виведення статистики
    print_statistics(gender_stats, age_stats, gender_age_stats)
    
    # Побудова діаграм
    print("=" * 70)
    print("Побудова діаграм...")
    print("=" * 70)
    
    plot_gender_distribution(gender_stats)
    plot_age_distribution(age_stats)
    plot_gender_by_age_categories(gender_age_stats)
    
    print()
    print("=" * 70)
    print("Аналіз завершено успішно!")
    print("Всі діаграми збережено у поточній директорії.")
    print("=" * 70)


if __name__ == "__main__":
    main()
