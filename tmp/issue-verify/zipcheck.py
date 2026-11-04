import struct, sys

path = sys.argv[1]
data = open(path, 'rb').read()
print(f"file size: {len(data)}")

# Find EOCD: scan backwards from EOF for the signature 0x06054b50.
# With trailing junk, the signature may not be in the last 22 bytes, so scan further.
sig = b'PK\x05\x06'
eocd = data.rfind(sig)
print(f"EOCD signature at offset: {eocd}")

if eocd < 0:
    print("no EOCD found"); sys.exit(1)

(eocd_sig, disk, cd_disk, n_disk, n_total, cd_size, cd_offset, comment_len) = struct.unpack('<IHHHHIIH', data[eocd:eocd+22])
print(f"entries this disk: {n_disk}, total: {n_total}")
print(f"cd_size: {cd_size}, cd_offset: {cd_offset}")
print(f"comment_len: {comment_len}")
record_end = eocd + 22 + comment_len
trailing = len(data) - record_end
print(f"EOCD record ends at: {record_end} (= {eocd} + 22 + {comment_len})")
print(f"bytes after EOCD record: {trailing}")
print(f"central dir consistent: cd_offset {cd_offset} + cd_size {cd_size} = {cd_offset + cd_size} (EOCD at {eocd}) -> {'YES' if cd_offset + cd_size == eocd else 'NO'}")

if trailing > 0:
    junk = data[record_end:]
    print(f"first 10 bytes hex: {junk[:10].hex(' ')}")
    try:
        text = junk.decode('utf-8', errors='replace')
        print("--- decoded trailing bytes (first 900 chars) ---")
        print(text[:900])
    except Exception as e:
        print("decode failed:", e)

# Also verify the fixed deleted-bytes variant: strip the trailing bytes, check zip integrity
if trailing > 0:
    stripped = data[:record_end]
    import zipfile, io
    try:
        zf = zipfile.ZipFile(io.BytesIO(stripped))
        names = zf.namelist()
        bad = zf.testzip()
        print(f"\nstripped archive: {len(names)} entries, testzip bad entry: {bad}")
        print("entries:", ", ".join(names))
    except Exception as e:
        print("stripped archive failed:", e)
