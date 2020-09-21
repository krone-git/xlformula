try:
    import tkinter as tk


    def _copy_to_clipboard(string):
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(string)
        root.update()
        root.destroy()

except (ImportError,) as e:
    def _copy_to_clipboard(string):
        pass
