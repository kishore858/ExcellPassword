import msoffcrypto
import io

# List of common passwords
passwords = ["123456", "password", "qwerty", "admin", "letmein", "welcome"]  
input_file = "excel_files/protected.xlsx"  # Use the folder path


with open(input_file, "rb") as file:
    encrypted = msoffcrypto.OfficeFile(file)

    for password in passwords:
        try:
            encrypted.load_key(password=password)
            decrypted_stream = io.BytesIO()
            encrypted.decrypt(decrypted_stream)
            print(f"✅ Password found: {password}")
            break
        except:
            print(f"❌ Tried password: {password}")
