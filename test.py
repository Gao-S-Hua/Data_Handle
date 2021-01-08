import psutil

for proc in psutil.process_iter():
    if proc.name() == "excel.exe":
        print("Excel is open")