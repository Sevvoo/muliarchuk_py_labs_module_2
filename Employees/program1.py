import csv
import random
from datetime import datetime
from faker import Faker

fake = Faker('uk_UA')

PATRONYMICS_MALE = [
    "Олександрович", "Іванович", "Петрович", "Миколайович", "Васильович",
    "Андрійович", "Сергійович", "Володимирович", "Дмитрович", "Юрійович",
    "Михайлович", "Богданович", "Ярославович", "Віталійович", "Анатолійович",
    "Григорович", "Романович", "Тарасович", "Павлович", "Степанович",
    "Вікторович", "Ігорович", "Максимович", "Леонідович", "Валерійович"
]

PATRONYMICS_FEMALE = [
    "Олександрівна", "Іванівна", "Петрівна", "Миколаївна", "Василівна",
    "Андріївна", "Сергіївна", "Володимирівна", "Дмитрівна", "Юріївна",
    "Михайлівна", "Богданівна", "Ярославівна", "Віталіївна", "Анатоліївна",
    "Григорівна", "Романівна", "Тарасівна", "Павлівна", "Степанівна",
    "Вікторівна", "Ігорівна", "Максимівна", "Леонідівна", "Валеріївна"
]


def generate_employee_record(gender):
    if gender == 'M':
        first_name = fake.first_name_male()
        last_name = fake.last_name_male()
        patronymic = random.choice(PATRONYMICS_MALE)
        gender_label = 'Ч'
    else:
        first_name = fake.first_name_female()
        last_name = fake.last_name_female()
        patronymic = random.choice(PATRONYMICS_FEMALE)
        gender_label = 'Ж'
    
    birth_date = fake.date_of_birth(minimum_age=17, maximum_age=87)
    
    # Генерація інших даних
    position = fake.job()
    city = fake.city()
    address = fake.address()
    phone = fake.phone_number()
    email = fake.email()
    
    return {
        'Прізвище': last_name,
        'Імя': first_name,
        'По бат': patronymic,
        'Стать': gender_label,
        'Дата народження': birth_date.strftime('%d.%m.%Y'),
        'Посада': position,
        'Місто проживання': city,
        'Адреса прож': address.replace('\n', ', '),
        'Телефон': phone,
        'Email': email
    }


def generate_csv_file(filename='employees.csv', num_records=500):

    genders = ['M'] * int(num_records * 0.6) + ['F'] * int(num_records * 0.4)
    random.shuffle(genders)
    
    fieldnames = [
        'Прізвище', 'Імя', 'По бат', 'Стать', 'Дата народження',
        'Посада', 'Місто проживання', 'Адреса прож', 'Телефон', 'Email'
    ]
    
    with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        
        for i, gender in enumerate(genders, 1):
            record = generate_employee_record(gender)
            writer.writerow(record)
            
            if i % 100 == 0:
                print(f"Згенеровано {i} записів...")
    
    print(f"\n✅ Успішно створено файл '{filename}' з {num_records} записами!")
    print(f"   - Чоловіків: {genders.count('M')} ({genders.count('M')/num_records*100:.1f}%)")
    print(f"   - Жінок: {genders.count('F')} ({genders.count('F')/num_records*100:.1f}%)")


def main():
    print("=" * 70)
    print("Програма 1 - Генерація CSV файлу з даними співробітників")
    print("=" * 70)
    print()
    
    generate_csv_file('employees.csv', 500)
    
    print("\nФайл збережено у поточній директорії.")


if __name__ == "__main__":
    main()
