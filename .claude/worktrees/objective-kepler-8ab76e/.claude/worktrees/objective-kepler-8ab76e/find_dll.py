import os
import psycopg2

# หาที่อยู่ของโฟลเดอร์ psycopg2
psycopg2_path = os.path.dirname(psycopg2.__file__)
print(f"Psycopg2 is located at: {psycopg2_path}")

# โดยปกติแล้ว ไฟล์ DLL จะอยู่ในโฟลเดอร์ .libs ที่อยู่ข้างใน
libs_path = os.path.join(psycopg2_path, '.libs')
print(f"Checking for DLLs in: {libs_path}")

if os.path.exists(libs_path):
    print("\n--- Found DLLs! ---")
    for filename in os.listdir(libs_path):
        if filename.endswith('.dll'):
            print(filename)
    print("--------------------")
else:
    print("\nCould not find a '.libs' folder. The DLLs might be in the main folder.")