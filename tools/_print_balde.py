import sys,traceback
sys.path.insert(0,r"C:\Users\USER\OneDrive\Desktop\Sistema_Helados")
out=r"C:\Users\USER\OneDrive\Desktop\Sistema_Helados\_print_result.txt"
try:
    import app
    be=app.make_backend()
    names=be.list_printers()
    pr=[n for n in names if "hprt" in n.lower() and "203" in n.lower()]
    cfg=app.load_config()
    p=pr[0]
    msg=""
    for name in ["Menta Granizado - Balde 6 kg"]:
        job=app.PrintJob(app.product_elements(name),100,80,1,cfg,"label")
        ok,text=be.status(p)
        msg+=f"{name}: status={ok},{text} -> "+str(be.send(p,job,"E-Karu balde "+name[:20]))+"\n"
    open(out,"w",encoding="utf-8").write(msg)
except Exception:
    open(out,"w",encoding="utf-8").write(traceback.format_exc())
