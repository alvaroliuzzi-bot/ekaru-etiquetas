import subprocess
cmd="Get-PrintJob -PrinterName 'One Driver' -ErrorAction SilentlyContinue | Format-List | Out-String; Get-Printer -Name 'One Driver' | Format-List Name,PrinterStatus,JobCount | Out-String; Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Operational' -MaxEvents 8 -ErrorAction SilentlyContinue | Select TimeCreated,Id,Message | Format-List | Out-String -Width 220"
r=subprocess.run(["powershell","-NoProfile","-Command",cmd],capture_output=True,text=True,encoding="cp850",errors="replace")
open(r"C:\Users\USER\OneDrive\Desktop\Sistema_Helados\_printers3.txt","w",encoding="utf-8").write(r.stdout+r.stderr)
