import sqlite3

# Veritabanı dosyasını oluştur ve bağlan
conn = sqlite3.connect('monthlyrobots.db')
cursor = conn.cursor()

# Tabloyu oluştur (eğer mevcut değilse)
cursor.execute('''
CREATE TABLE IF NOT EXISTS monthlyrobots (
    robot_name TEXT,
    date DATETIME,
    percentage TEXT
)
''')

# Veri girişi
data = [
    ('ROBOT', '27/01/2024', '%62.28'),
    ('ROBOT2', '27/01/2024', '%66.08'),
    ('ROBOT3', '27/01/2024', '%34.06'),
    ('ROBOT', '03/02/2024', '%52.28'),
    ('ROBOT2', '03/02/2024', '%86.08'),
    ('ROBOT3', '03/02/2024', '%14.06'),
    ('ROBOT', '10/02/2024', '%72.28'),
    ('ROBOT2', '10/02/2024', '%26.08'),
    ('ROBOT3', '10/02/2024', '%44.06')
]

# Verileri tabloya ekle
cursor.executemany('INSERT INTO monthlyrobots (robot_name, date, percentage) VALUES (?, ?, ?)', data)

# Değişiklikleri kaydet
conn.commit()

# Bağlantıyı kapat
conn.close()

print("Veriler başarıyla kaydedildi.")
