import subprocess
cmd="Get-PnpDevice -PresentOnly | Where-Object {$_.Class -in 'USB','Printer','PrintQueue'} | Select Class,FriendlyName,Status | Format-Table -Auto | Out-String -Width 200; Get-PrinterProperty -PrinterName 'One Driver' | Select PropertyName,Value | Format-Table -Auto | Out-String -Width 200; Get-PrintConfiguration -PrinterName 'One Driver' | Format-List | Out-String"
r=subprocess.run(["powershell","-NoProfile","-Command",cmd],capture_output=True,text=True,encoding="cp850",errors="replace")
open(r"C:\Users\USER\OneDrive\Desktop\Sistema_Helados\_printers2.txt","w",encoding="utf-8").write(r.stdout+r.stderr)
