import subprocess
r=subprocess.run(["powershell","-NoProfile","-Command","Get-Printer | Select Name,DriverName,PortName,PrinterStatus | Format-List; Get-PrintConfiguration -PrinterName '*' | Select PrinterName,PaperSize | Format-List"],capture_output=True,text=True,encoding="cp850",errors="replace")
open(r"C:\Users\USER\OneDrive\Desktop\Sistema_Helados\_printers.txt","w",encoding="utf-8").write(r.stdout+r.stderr)
