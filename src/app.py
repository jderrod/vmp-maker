import tkinter as tk

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Visual Manufacturing Procedures")
        self.geometry("1200x800")

        # Container to hold all frames
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        # Initialize all frames here
        # for F in (HomePage, EditorPage): # We will create these frames next
        #     frame = F(container, self)
        #     self.frames[F] = frame
        #     frame.grid(row=0, column=0, sticky="nsew")

        # self.show_frame(HomePage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()
