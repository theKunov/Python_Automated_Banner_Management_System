import threading, string
import tkinter as tk
from tkinter import ttk
  

def banner_setup():
    from setup_main import run_prog
    run_prog()

def screenshot_maker():
    from screenshot_main import screenshot
    screenshot()

def move():
    from move_to_directory import move_to_dir
    move_to_dir()

LARGEFONT =("Verdana", 15)
COLOR = ("red")
  
class tkinterApp(tk.Tk):
     
    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
         
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
  
        # initializing frames to an empty array
        self.frames = {} 
  
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (Setup, Screenshot, Move):
  
            frame = F(container, self)
  
            # initializing frame of that object from
            # startpage, Screenshot, Move respectively with
            # for loop
            self.frames[F] = frame
  
            frame.grid(row = 0, column = 0, sticky ="nsew")
  
        self.show_frame(Setup)
  
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()
  
# first window frame startpage
  
class Setup(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        button0 = ttk.Button(self, text ="Setup",
        command = lambda : controller.show_frame(Setup))
     
        # putting the button in its place by
        # using grid
        button0.grid(row = 0, column = 0, padx= 10, pady = 10)


        button1 = ttk.Button(self, text ="Screenshots",
        command = lambda : controller.show_frame(Screenshot))
     
        # putting the button in its place by
        # using grid
        button1.grid(row = 0, column = 1, pady = 10)
  
        ## button to show frame 2 with text layout2
        button2 = ttk.Button(self, text ="Copy to directory",
        command = lambda : controller.show_frame(Move))
     
        # putting the button in its place by
        # using grid
        button2.grid(row = 0, column = 2, pady = 10)
         

        # label of frame Layout 2
        label = ttk.Label(self, text ="Banner Setup", font = LARGEFONT)
        # putting the grid in its place by using
        # grid
        label.grid(row = 1, column = 1, padx = 10, pady = 10)


        button3 = ttk.Button(self, text ="Start",
        command = lambda:threading.Thread(target=banner_setup).start() )
        button3.grid(row = 2, column = 1, padx = 10, pady = 10)
         
        
  
          
  
  
# second window frame Screenshot
class Screenshot(tk.Frame):
     
    def __init__(self, parent, controller):
         
        tk.Frame.__init__(self, parent)
        
        button0 = ttk.Button(self, text ="Setup",
        command = lambda : controller.show_frame(Setup))
     
        # putting the button in its place by
        # using grid
        button0.grid(row = 0, column = 0, padx= 10, pady = 10)


        button1 = ttk.Button(self, text ="Screenshots",
        command = lambda : controller.show_frame(Screenshot))
     
        # putting the button in its place by
        # using grid
        button1.grid(row = 0, column = 1, pady = 10)
  
        ## button to show frame 2 with text layout2
        button2 = ttk.Button(self, text ="Copy to directory",
        command = lambda : controller.show_frame(Move))
     
        # putting the button in its place by
        # using grid
        button2.grid(row = 0, column = 2, pady = 10)
         

        # label of frame Layout 2
        label = ttk.Label(self, text ="Screenshots", font = LARGEFONT)
        # putting the grid in its place by using
        # grid
        label.grid(row = 1, column = 1, padx = 10, pady = 10)

        button3 = ttk.Button(self, text ="Start",
        command = lambda:threading.Thread(target=screenshot_maker).start() )
        button3.grid(row = 2, column = 1, padx = 10, pady = 10)


        # label of frame Layout 2
        # part = "3"
        # label[part] = ttk.Label(self, text ="Test")
        # putting the grid in its place by using
        # grid
        # label[part].grid(row = part, column = 0, padx = 10, pady = 10)
  
  
  
  
# third window frame Move
class Move(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        
        button0 = ttk.Button(self, text ="Setup",
        command = lambda : controller.show_frame(Setup))
     
        # putting the button in its place by
        # using grid
        button0.grid(row = 0, column = 0, padx= 10, pady = 10)


        button1 = ttk.Button(self, text ="Screenshots",
        command = lambda : controller.show_frame(Screenshot))
     
        # putting the button in its place by
        # using grid
        button1.grid(row = 0, column = 1, pady = 10)
  
        ## button to show frame 2 with text layout2
        button2 = ttk.Button(self, text ="Copy to directory",
        command = lambda : controller.show_frame(Move))
     
        # putting the button in its place by
        # using grid
        button2.grid(row = 0, column = 2, pady = 10)
         

        # label of frame Layout 2
        label = ttk.Label(self, text ="Copy to Directory", font = LARGEFONT)
        # putting the grid in its place by using
        # grid
        label.grid(row = 1, column = 1, padx = 10, pady = 10)


        button3 = ttk.Button(self, text ="Start",
        command = lambda:threading.Thread(target=move).start() )
        button3.grid(row = 2, column = 1, padx = 10, pady = 10)
  
  
# Driver Code
app = tkinterApp()
app.title('Automated Banner Manager')
app.mainloop()