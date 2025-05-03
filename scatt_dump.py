import olefile, binascii, textwrap, sys

if len(sys.argv) < 2:
    print("python scatt_dump.py <datei.scatt>")
    sys.exit(1)

with olefile.OleFileIO(sys.argv[1]) as ole:
    data = ole.openstream("Contents").read(512)   # erste 512 Bytes
    hexstr = binascii.hexlify(data).decode()

# hübsch 16 Bytes pro Zeile ausgeben
for i in range(0, len(hexstr), 32):
    chunk = hexstr[i:i+32]
    print(f"{i//2:04x}: {' '.join(textwrap.wrap(chunk, 2))}")
