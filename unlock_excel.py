import msoffcrypto
import pandas as pd
import io

# File paths
input_file = "excel_files/protected.xlsx"  # Change this to your actual file
output_file = "unlocked.xlsx"

# Enter the password (if you remember it)
password = "yourpassword"  # Replace with your password

# Open the encrypted Excel file
with open(input_file, "rb") as file:
    encrypted = msoffcrypto.OfficeFile(file)
    encrypted.load_key(password=password)  # Load key using password
    
    decrypted_stream = io.BytesIO()
    encrypted.decrypt(decrypted_stream)

# Read and save the unlocked Excel file
df = pd.read_excel(decrypted_stream, engine="openpyxl")
df.to_excel(output_file, index=False)

print(f"✅ Excel file unlocked and saved as '{output_file}'")
