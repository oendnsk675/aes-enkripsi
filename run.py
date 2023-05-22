from utils import * 
import subprocess, time
from tqdm import tqdm

# Blok teks
plain = [
    [0x10, 0x20, 0xea, 0x19],
    [0x04, 0x54, 0xae, 0x5d],
    [0x3d, 0xed, 0x76, 0x0f],
    [0x8a, 0x3d, 0xaf, 0x20]
]

plain_tmp = [
    [0x10, 0x20, 0xea, 0x19],
    [0x04, 0x54, 0xae, 0x5d],
    [0x3d, 0xed, 0x76, 0x0f],
    [0x8a, 0x3d, 0xaf, 0x20]
]

constant_matrix = [
    [0x02, 0x03, 0x01, 0x01],
    [0x01, 0x02, 0x03, 0x01],
    [0x01, 0x01, 0x02, 0x03],
    [0x03, 0x01, 0x01, 0x02]
]


# Kunci utama
key = [
    [0x16, 0x36, 0x88, 0x3c],
    [0x15, 0x2d, 0x15, 0x4f],
    [0x7e, 0xaf, 0x47, 0xfa],
    [0x2b, 0x28, 0xab, 0x3c]
]

filename = "ENC.xlsx"

# Mendapatkan daftar proses yang menggunakan file
# UNCOMMENT Hanya untuk debuging saja
# if os.path.exists(filename):
#     try:
#         process_name = "EXCEL.EXE"
#         subprocess.run(['taskkill', '/F', '/IM', process_name], check=True)
#         print(f"Proses '{process_name}' berhasil dihentikan.")
#         time.sleep(2)
#         os.remove(filename)
#     except subprocess.CalledProcessError as e:
#         print(f"Terjadi kesalahan saat menghentikan proses '{process_name}': {e}")

#------------------- key schadule ------------------------- 

# Membuat key schedule
r_key_schedule = key_schedule(key)

# write on excel

start_row = 6
start_column = 2
start_row_tmp = start_row

draw_table(row=3, col=2, value="PENENTUAN KEY SCHADULE SEBANYAK 10 KEY TIDAK TERMASUK KEY INITIAL(0)", filename=filename,border_none=False)
print("[1] KEY SCHADULE")
for n, ks in enumerate(r_key_schedule):
        if(n == 0):
            draw_table(start_row_tmp - 1, start_column + 1, f"K({n}) / Chiper key", filename, False)
        else:
            draw_table(start_row_tmp - 1, start_column + 1, f"K({n})", filename, False)

        for row in range(len(ks)):
            for col in range(len(ks[row])):
                value = hex(int(ks[row][col], 16))[2:].upper()
                draw_table(start_row_tmp + row, start_column + col, value, filename)

        start_row_tmp += (start_row + 1)

# ----------------- initial round / add round key ------------------

print("[2] INITIAL ROUND")
draw_table(row=83, col=8, value="INITIAL ROUND", filename=filename,border_none=False)
r_add_round_key = add_round_key(plain_tmp, r_key_schedule[0])

for row in range(4):
    for col in range(4):
        p = int(plain[row][col], 16) if isinstance(plain[row][col], str) else plain[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=86 + row, col=2 + col, value=value, filename=filename)

draw_table(row=87, col=6, value="XOR", filename=filename,border_none=False)

for row in range(4):
    for col in range(4):
        p = int(r_key_schedule[0][row][col], 16) if isinstance(r_key_schedule[0][row][col], str) else r_key_schedule[0][row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=86 + row, col=7 + col, value=value, filename=filename)

draw_table(row=87, col=11, value="=", filename=filename,border_none=False)
for row in range(4):
    for col in range(4):
        p = int(r_add_round_key[row][col], 16) if isinstance(r_add_round_key[row][col], str) else r_add_round_key[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=86 + row, col=12 + col, value=value, filename=filename)
# ----------------- main round -------------------------------------

draw_table(row=92, col=8, value="Main Round Has 9 Round", filename=filename,border_none=False)

print("[3] MAIN ROUND")

for n in tqdm(range(1, 10), desc="Progress", unit="round" , bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}"):
    # print(f"Round - {n}", end="\r", flush=True)
    draw_table(row=95 + (30 * (n-1)), col=8, value=f"ROUND {n}", filename=filename,border_none=False)

    # /------------------------------ SUB BYTES -------------------------------------------
    draw_table(row=97 + (30 * (n-1)), col=8, value=f"SUB BYTES", filename=filename,border_none=False)
    for row in range(4):
        for col in range(4):
            p = int(r_add_round_key[row][col], 16) if isinstance(r_add_round_key[row][col], str) else r_add_round_key[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + (7*(n-1) + (23 * (n-1))), col=2 + col, value=value, filename=filename)

    r_sub_bytes = sub_bytes(r_add_round_key)
  
    draw_table(row=98 + (7*(n-1)) + (23 * (n-1)), col=7, value=f"Hasil Dari Subtitusi Dengan S-BOX =>", filename=filename,border_none=False)
    
    for row in range(4):
        for col in range(4):
            p = int(r_sub_bytes[row][col], 16) if isinstance(r_sub_bytes[row][col], str) else r_sub_bytes[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + (7*(n-1)) + (23 * (n-1)), col=12 + col, value=value, filename=filename)
    # /------------------------------ /SUB BYTES -------------------------------------------

    # /------------------------------ SHIFT ROWS -------------------------------------------
    draw_table(row=97 + (7*n) + (23 * (n-1)), col=8, value=f"SHIFT ROWS", filename=filename,border_none=False)
    for row in range(4):
        for col in range(4):
            p = int(r_sub_bytes[row][col], 16) if isinstance(r_sub_bytes[row][col], str) else r_sub_bytes[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + (7*n) + (23 * (n-1)), col=2 + col, value=value, filename=filename)

    r_shift_rows = shift_rows(r_sub_bytes)

    draw_table(row=98 + (7*n) + (23 * (n-1)), col=7, value=f"Hasil Dari SHIFT ROWS =>", filename=filename,border_none=False)

    for row in range(4):
        for col in range(4):
            p = int(r_shift_rows[row][col], 16) if isinstance(r_shift_rows[row][col], str) else r_shift_rows[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + (7*n) + (23 * (n-1)), col=12 + col, value=value, filename=filename)
    # /------------------------------ /SHIFT ROWS -------------------------------------------
    
    # /------------------------------ MIX COLUMNS -------------------------------------------
    draw_table(row=97 + ((7*n) + 7) + (23 * (n-1)), col=8, value=f"MIX COLUMNS", filename=filename,border_none=False)
    for row in range(4):
        for col in range(4):
            p = int(r_shift_rows[row][col], 16) if isinstance(r_shift_rows[row][col], str) else r_shift_rows[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 7) + (23 * (n-1)), col=2 + col, value=value, filename=filename)

    r_mix_columns = mix_columns(r_shift_rows)

    draw_table(row=98 + ((7*n) + 7) + (23 * (n-1)), col=7, value=f"Hasil Dari MIX COLUMNS =>", filename=filename,border_none=False)

    draw_table(row=100 + ((7*n) + 7) + (23 * (n-1)), col=6, value=f".", filename=filename,border_none=False)

    for row in range(4):
        for col in range(4):
            p = int(constant_matrix[row][col], 16) if isinstance(constant_matrix[row][col], str) else constant_matrix[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 7) + (23 * (n-1)), col=7 + col, value=value, filename=filename)

    draw_table(row=100 + ((7*n) + 7) + (23 * (n-1)), col=11, value=f"=", filename=filename,border_none=False)

    for row in range(4):
        for col in range(4):
            p = int(r_mix_columns[row][col], 16) if isinstance(r_mix_columns[row][col], str) else r_mix_columns[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 7) + (23 * (n-1)), col=12 + col, value=value, filename=filename)
    # /------------------------------ /MIX COLUMNS -------------------------------------------
    
    # /------------------------------ ADD ROUND KEY -------------------------------------------
    draw_table(row=97 + ((7*n) + 14) + (23 * (n-1)), col=8, value=f"ADD ROUND KEY", filename=filename,border_none=False)
    for row in range(4):
        for col in range(4):
            p = int(r_mix_columns[row][col], 16) if isinstance(r_mix_columns[row][col], str) else r_mix_columns[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 14) + (23 * (n-1)), col=2 + col, value=value, filename=filename)

    draw_table(row=100 + ((7*n) + 14) + (23 * (n-1)), col=6, value=f"XOR", filename=filename,border_none=False)

    for row in range(4):
        for col in range(4):
            p = int(r_key_schedule[n][row][col], 16) if isinstance(r_key_schedule[n][row][col], str) else r_key_schedule[n][row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 14) + (23 * (n-1)), col=7 + col, value=value, filename=filename)

    r_add_round_key = add_round_key(r_mix_columns, r_key_schedule[n])

    draw_table(row=98 + ((7*n) + 14) + (23 * (n-1)), col=7, value=f"Hasil Dari ADD ROUND KEY =>", filename=filename,border_none=False)

    draw_table(row=100 + ((7*n) + 14) + (23 * (n-1)), col=11, value=f"=", filename=filename,border_none=False)

    for row in range(4):
        for col in range(4):
            p = int(r_add_round_key[row][col], 16) if isinstance(r_add_round_key[row][col], str) else r_add_round_key[row][col]
            value = hex(p)[2:].upper().zfill(2)
            draw_table(row=(99 + row) + ((7*n) + 14) + (23 * (n-1)), col=12 + col, value=value, filename=filename)
    # /------------------------------ /ADD ROUND KEY -------------------------------------------
    # if n == 5:
    #     exit()
    # continue

print("[4] FINAL ROUND")

draw_table(row=366, col=8, value="FINAL ROUND(10)", filename=filename,border_none=False)
# /------------------------------ SUB BYTES -------------------------------------------
draw_table(row=368, col=8, value="SUB BYTES", filename=filename,border_none=False)
for row in range(4):
    for col in range(4):
        p = int(r_add_round_key[row][col], 16) if isinstance(r_add_round_key[row][col], str) else r_add_round_key[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(370 + row), col=2 + col, value=value, filename=filename)
r_sub_bytes = sub_bytes(r_add_round_key)

draw_table(row=369, col=7, value=f"Hasil Dari Subtitusi Dengan S-BOX =>", filename=filename,border_none=False)

for row in range(4):
    for col in range(4):
        p = int(r_sub_bytes[row][col], 16) if isinstance(r_sub_bytes[row][col], str) else r_sub_bytes[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(370 + row), col=12 + col, value=value, filename=filename)
# /------------------------------ /SUB BYTES -------------------------------------------

# /------------------------------ SHIFT ROWS -------------------------------------------
draw_table(row=375, col=8, value="SHIFT ROWS", filename=filename,border_none=False)
for row in range(4):
    for col in range(4):
        p = int(r_sub_bytes[row][col], 16) if isinstance(r_sub_bytes[row][col], str) else r_sub_bytes[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(377 + row), col=2 + col, value=value, filename=filename)
r_shift_rows = shift_rows(r_sub_bytes)

draw_table(row=376, col=7, value=f"Hasil Dari Subtitusi Dengan SHIFT ROWS =>", filename=filename,border_none=False)

for row in range(4):
    for col in range(4):
        p = int(r_shift_rows[row][col], 16) if isinstance(r_shift_rows[row][col], str) else r_shift_rows[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(377 + row), col=12 + col, value=value, filename=filename)
# /------------------------------ /SHIFT ROWS -------------------------------------------

# /------------------------------ ADD ROUND KEY -------------------------------------------
draw_table(row=382, col=8, value="SHIFT ROWS", filename=filename,border_none=False)
for row in range(4):
    for col in range(4):
        p = int(r_shift_rows[row][col], 16) if isinstance(r_sub_bytes[row][col], str) else r_sub_bytes[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(384 + row), col=2 + col, value=value, filename=filename)

r_add_round_key = add_round_key(r_shift_rows, r_key_schedule[-1])

draw_table(row=383, col=7, value=f"Hasil Dari Subtitusi Dengan SHIFT ROWS =>", filename=filename,border_none=False)

for row in range(4):
    for col in range(4):
        p = int(r_add_round_key[row][col], 16) if isinstance(r_add_round_key[row][col], str) else r_add_round_key[row][col]
        value = hex(p)[2:].upper().zfill(2)
        draw_table(row=(384 + row), col=12 + col, value=value, filename=filename)
# /------------------------------ /ADD ROUND KEY -------------------------------------------

draw_table(row=390, col=7, value=f"JADI HASIL ENKRIPSI", filename=filename,border_none=False)

result = ""
for col in range(len(r_add_round_key[0])):
    for row in range(len(r_add_round_key)):
        element = r_add_round_key[row][col]
        value = element[2:].zfill(2)  # Menghilangkan '0x' dan menambahkan awalan '0' jika diperlukan
        result += value

draw_table(row=391, col=7, value=result, filename=filename,border_none=False)
print(f"Hasil Enkripsi : {result}")
