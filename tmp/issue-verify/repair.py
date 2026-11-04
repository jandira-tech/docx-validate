import sys
src, dst = sys.argv[1], sys.argv[2]
data = open(src, 'rb').read()
eocd = data.rfind(b'PK\x05\x06')
comment_len = int.from_bytes(data[eocd+20:eocd+22], 'little')
end = eocd + 22 + comment_len
open(dst, 'wb').write(data[:end])
print(f"kept {end} bytes (dropped {len(data)-end})")
