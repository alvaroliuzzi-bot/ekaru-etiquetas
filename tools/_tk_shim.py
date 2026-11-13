# -*- coding: utf-8 -*-
"""Sustituto mínimo de tkinter y customtkinter para PROBAR app.py sin pantalla (Linux, servidores, CI).

No dibuja nada: registra las llamadas, comprueba que no se mezclen pack() y grid() en un mismo contenedor
(Tk lo rechaza) y deja ejecutar los eventos pendientes con pump(). En la PC de Glacé, con tkinter y
customtkinter instalados, no hace falta: install() solo actúa si faltan.
"""
import sys
import types

PENDING = []          # (ms, función, args) de after()/after_idle()
UNKNOWN_CALLS = set()  # métodos de widgets que el sustituto no conoce (para revisar a mano)
CREATED = []          # todos los widgets creados
ANSWERS = {}          # respuestas programadas de los cuadros de diálogo: nombre -> valor (o lista)
DIALOG_CALLS = []     # (nombre, título, texto)
STATE = {"focus": None}


class TclError(Exception):
    pass


def _answer(name, default):
    value = ANSWERS.get(name, default)
    if isinstance(value, list):
        return value.pop(0) if value else default
    return value


class Var:
    def __init__(self, master=None, value="", name=None):
        self._value = value
        self._traces = []

    def get(self):
        return self._value

    def set(self, value):
        self._value = value
        for callback in list(self._traces):
            callback("var", "", "write")

    def trace_add(self, mode, callback):
        self._traces.append(callback)
        return "trace"


class Widget:
    def __init__(self, master=None, *args, **kw):
        self.master = master
        self.kw = dict(kw)
        self.children = []
        self._gm = None
        self._used = set()
        self._bindings = {}
        self._alive = True
        self._text = ""
        self._items = []
        if master is not None and hasattr(master, "children"):
            master.children.append(self)
        CREATED.append(self)

    # -- geometría: Tk no permite mezclar pack y grid dentro de un mismo contenedor
    def _manage(self, kind):
        parent = self.master
        if parent is not None and kind in ("pack", "grid"):
            other = {"pack", "grid"} - {kind}
            if parent._used & other:
                raise TclError(f"cannot use geometry manager {kind} inside {type(parent).__name__} "
                               f"which already has slaves managed by {sorted(parent._used)}")
            parent._used.add(kind)
        self._gm = kind

    def pack(self, **kw):
        self._manage("pack")

    def grid(self, **kw):
        self._manage("grid")

    def place(self, **kw):
        self._manage("place")

    def pack_forget(self):
        self._gm = None

    def grid_forget(self):
        self._gm = None

    def grid_remove(self):
        self._gm = None

    def grid_propagate(self, *a):
        pass

    def pack_propagate(self, *a):
        pass

    def grid_columnconfigure(self, *a, **kw):
        pass

    def grid_rowconfigure(self, *a, **kw):
        pass

    # -- configuración
    def configure(self, **kw):
        self.kw.update(kw)
        if "text" in kw:
            self._text = kw["text"]

    config = configure

    def cget(self, key):
        return self.kw.get(key)

    def bind(self, sequence, func, add=None):
        self._bindings[sequence] = func

    def after(self, ms, func=None, *args):
        if func is not None:
            PENDING.append((ms, func, args))
        return f"after#{len(PENDING)}"

    def after_idle(self, func, *args):
        PENDING.append((0, func, args))
        return f"idle#{len(PENDING)}"

    def after_cancel(self, ident):
        pass

    def winfo_width(self):
        return int(self.kw.get("width", 0) or 900)

    def winfo_height(self):
        return int(self.kw.get("height", 0) or 600)

    def winfo_screenwidth(self):
        return 1920

    def winfo_screenheight(self):
        return 1080

    def winfo_exists(self):
        return 1 if self._alive else 0

    def focus_set(self):
        STATE["focus"] = self

    focus = focus_set

    def destroy(self):
        self._alive = False
        for child in list(self.children):
            child.destroy()

    def lift(self, *a):
        pass

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        UNKNOWN_CALLS.add(f"{type(self).__name__}.{name}")
        return lambda *a, **k: None


class Frame(Widget):
    pass


class Label(Widget):
    pass


class Button(Widget):
    def invoke(self):
        command = self.kw.get("command")
        if command:
            return command()


class Entry(Widget):
    def __init__(self, master=None, *a, **kw):
        super().__init__(master, *a, **kw)
        self._text = ""

    def _var(self):
        return self.kw.get("textvariable")

    def get(self):
        var = self._var()
        return var.get() if var is not None else self._text

    def insert(self, index, text):
        var = self._var()
        if var is not None:
            var.set(var.get() + str(text))
        else:
            self._text += str(text)

    def delete(self, a, b=None):
        var = self._var()
        if var is not None:
            var.set("")
        else:
            self._text = ""

    def select_range(self, a, b):
        pass


class Textbox(Widget):
    def __init__(self, master=None, *a, **kw):
        super().__init__(master, *a, **kw)
        self._text = ""

    def insert(self, index, text, *tags):
        if index in ("1.0", "0.0"):
            self._text = str(text) + self._text
        else:
            self._text += str(text)

    def get(self, a="1.0", b="end"):
        return self._text

    def delete(self, a, b=None):
        self._text = ""


class OptionMenu(Widget):
    def __init__(self, master=None, *a, **kw):
        super().__init__(master, *a, **kw)
        self._value = kw.get("values", [""])[0] if kw.get("values") else ""
        var = kw.get("variable")
        if var is not None:
            self._value = var.get()

    def set(self, value):
        self._value = value
        var = self.kw.get("variable")
        if var is not None:
            var.set(value)

    def get(self):
        var = self.kw.get("variable")
        return var.get() if var is not None else self._value

    def choose(self, value):
        """Como si la persona eligiera «value» en el menú."""
        self.set(value)
        command = self.kw.get("command")
        if command:
            command(value)


class SegmentedButton(OptionMenu):
    pass


class Switch(Widget):
    def __init__(self, master=None, *a, **kw):
        super().__init__(master, *a, **kw)
        self._on = False

    def select(self):
        self._on = True

    def deselect(self):
        self._on = False

    def get(self):
        return 1 if self._on else 0


class Canvas(Widget):
    def __init__(self, master=None, *a, **kw):
        super().__init__(master, *a, **kw)
        self._next = 0
        self._ids = []
        self.tags = {}

    def _new(self, kind, *args, **kw):
        self._next += 1
        self._ids.append(self._next)
        self._items.append((self._next, kind, args, kw))
        return self._next

    def create_rectangle(self, *a, **kw):
        return self._new("rect", *a, **kw)

    def create_text(self, *a, **kw):
        return self._new("text", *a, **kw)

    def create_line(self, *a, **kw):
        return self._new("line", *a, **kw)

    def create_image(self, *a, **kw):
        return self._new("image", *a, **kw)

    def create_window(self, *a, **kw):
        return self._new("window", *a, **kw)

    def delete(self, *tags):
        if "all" in tags:
            self._ids, self._items = [], []
        else:
            self._items = [i for i in self._items if i[3].get("tags") not in tags and f"row{i[0]}" not in tags]
            self._ids = [i[0] for i in self._items]

    def find_all(self):
        return tuple(self._ids)

    def addtag_withtag(self, tag, item):
        self.tags.setdefault(tag, []).append(item)

    def move(self, *a):
        pass

    def itemconfigure(self, *a, **kw):
        pass


class PhotoImage:
    def __init__(self, *a, **kw):
        self.kw = kw

    def put(self, *a, **kw):
        pass


class Font:
    def __init__(self, *a, **kw):
        self.kw = kw

    def measure(self, text):
        return len(str(text)) * 7


class CTkFont(Font):
    pass


class CTk(Widget):
    def __init__(self, *a, **kw):
        super().__init__(None, *a, **kw)
        self.destroyed = False

    def _get_window_scaling(self):
        return 1.0

    def title(self, *a):
        pass

    def geometry(self, *a):
        pass

    def minsize(self, *a):
        pass

    def state(self, *a):
        pass

    def protocol(self, *a):
        pass

    def iconbitmap(self, *a):
        pass

    def mainloop(self):
        pass

    def clipboard_clear(self):
        self.clipboard = ""

    def clipboard_append(self, text):
        self.clipboard = getattr(self, "clipboard", "") + text


class Toplevel(CTk):
    def __init__(self, master=None, *a, **kw):
        Widget.__init__(self, master, *a, **kw)
        self.destroyed = False

    def transient(self, *a):
        pass

    def resizable(self, *a):
        pass

    def grab_set(self):
        pass


def pump(rounds=3):
    """Ejecuta los after()/after_idle() pendientes (una pasada por ronda; los que se reprograman esperan)."""
    for _ in range(rounds):
        batch = list(PENDING)
        del PENDING[:]
        for _ms, func, args in batch:
            func(*args)


def install(force=False):
    """Registra los módulos falsos si tkinter/customtkinter no están (o si force=True)."""
    if not force:
        try:
            import customtkinter  # noqa: F401
            import tkinter  # noqa: F401
            return False
        except ImportError:
            pass
    tk = types.ModuleType("tkinter")
    tk.TclError = TclError
    tk.StringVar = Var
    tk.IntVar = Var
    tk.BooleanVar = Var
    tk.Canvas = Canvas
    tk.Entry = Entry
    tk.Frame = Frame
    tk.Label = Label
    tk.PhotoImage = PhotoImage
    tk.Tk = CTk
    mb = types.ModuleType("tkinter.messagebox")
    fd = types.ModuleType("tkinter.filedialog")
    sd = types.ModuleType("tkinter.simpledialog")
    tf = types.ModuleType("tkinter.font")

    def dialog(name, default):
        def run(title=None, message=None, **kw):
            DIALOG_CALLS.append((name, title, message))
            return _answer(name, default)
        return run

    for name, default in (("showinfo", "ok"), ("showwarning", "ok"), ("showerror", "ok"), ("askyesno", True),
                          ("askokcancel", True), ("askyesnocancel", True)):
        setattr(mb, name, dialog(name, default))
    for name, default in (("asksaveasfilename", ""), ("askopenfilename", ""), ("askdirectory", "")):
        setattr(fd, name, dialog(name, default))
    for name, default in (("askstring", None), ("askinteger", None)):
        setattr(sd, name, dialog(name, default))
    tf.Font = Font
    tf.families = lambda *a, **k: ("Consolas", "Segoe UI")
    tk.messagebox, tk.filedialog, tk.simpledialog, tk.font = mb, fd, sd, tf
    ctk = types.ModuleType("customtkinter")
    ctk.CTk = CTk
    ctk.CTkToplevel = Toplevel
    ctk.CTkFrame = Frame
    ctk.CTkLabel = Label
    ctk.CTkButton = Button
    ctk.CTkEntry = Entry
    ctk.CTkTextbox = Textbox
    ctk.CTkOptionMenu = OptionMenu
    ctk.CTkSegmentedButton = SegmentedButton
    ctk.CTkSwitch = Switch
    ctk.CTkScrollableFrame = Frame
    ctk.CTkScrollbar = type("CTkScrollbar", (Widget,), {"set": lambda self, *a: None})
    ctk.CTkFont = CTkFont
    ctk.set_appearance_mode = lambda *a, **k: None
    for module in (tk, mb, fd, sd, tf, ctk):
        sys.modules[module.__name__] = module
    return True
