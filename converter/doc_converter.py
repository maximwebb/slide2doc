def process_files(fs):
    res = ""
    for f in fs:
        with open("./doc.txt", 'wb+') as dst:
            lines = f.read()
            dst.write(lines)
            res += lines.decode("utf-8")
    return res