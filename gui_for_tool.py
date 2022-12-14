from tkinter import * 
# creat new window 
windown = Tk()
windown.title("Geanerate Report MPT")
windown.geometry("350x200")
# handle when click button
def click_button():
    """
    handle infor in function when click  
    """
    lbl.configure(text="BUTTON IS CLICKED")
# Doing button 
lbl = Label(windown,text="Hello")
lbl.grid(column=0,row=0)
button = Button(windown,text="Click Me",bg="orange",fg="red",command=click_button)
#thiet lap vi tri mau nen 
button.grid(column=1, row=0)

windown.mainloop()