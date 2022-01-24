from cgitb import text
from email.mime import image
from tkinter import * 
from tkinter import filedialog
from turtle import title

raiz = Tk()

raiz.title('Asignaciones Edenor')
raiz.iconbitmap('Apps/excel.ico')
#raiz.geometry('300x200')
raiz.config(bg='#4F8CDC') #Celeste

#Frames:
miFrame = Frame()
miFrame.pack()
miFrame.config(bg='#71032E', width='500', height='200') #Rojo

#Labels (Widgets): Para mostrar texto o imágenes.
#No se puede interactuar con él
miLabel = Label()
miLabel = Label(raiz, text='Texto de prueba').place(x=200, y=200)

miImagen = PhotoImage(file='Apps/edenor_logo.png')

Label(miFrame, image=miImagen).place(x=10, y=10)


def abrirArchivo():
    archivo = filedialog.askopenfilename(title='abrir', initialdir='C:/')
    print(archivo)

Button(raiz, text='Abrir Archivo', command=abrirArchivo).pack()


raiz.mainloop()