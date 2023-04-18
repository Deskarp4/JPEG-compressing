from tkinter import *
from tkinter import filedialog
import tkinter.ttk as ttk
from PIL import Image
from win32com.shell import shell, shellcon
import os




# Window
window = Tk()
window.title('Image compressing')
window.geometry('700x300')

# Button
def clicked():
    global file_directory
    global im

    file_directory = filedialog.askopenfilename(filetypes=[('Images', '*.png; *.jpg; *.gif')])
    final_name = '/Compr_'+((((file_directory.split('/'))[-1:])[0])[:-4])+'.jpg'
    desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0) + final_name
    save_directory = desktop.replace('\\', '/')

    btn.destroy()
    
    resize_value = 1
    im = Image.open(file_directory)
    im = im.convert('RGB')
    im = im.resize((int(im.size[0] * resize_value), int(im.size[1] * resize_value)), Image.Resampling.BILINEAR)

    start_size = (os.stat(file_directory).st_size)//1000

    if start_size < 10: pads = 0
    if start_size < 100 and start_size >= 10: pads = 5
    if start_size < 1000 and start_size >= 100: pads = 10
    if start_size >= 1000: pads = 15

    lbl_start = Label(window, text='Изначальный размер - ' + str(start_size) +' килобайт', fg = 'black', font=("Arial", 15), width=32)  
    lbl_start.grid(column=0, row=1, stick='w', pady=5, padx=pads)  

    im.save(save_directory, quality = 0, optimize=True)
    min_size = (os.stat(save_directory).st_size//1000)

    if min_size < 10: padn = 6
    if min_size < 100 and min_size >= 10: padn = 11
    if min_size < 1000 and min_size >= 100: padn = 16
    if min_size >= 1000: padn = 21

    lbl_min = Label(text='Минимально возможный размер - ' + str(min_size) + ' килобайт',fg = 'black', font=("Arial", 15), width=40)
    lbl_min.grid(column=0, row=2, stick='w', pady=5, padx=padn)  

    lbl_need = Label(text=' Введите желаемый размер в килобайтах: ', fg = 'black', font=("Arial", 15), width=36)
    lbl_need.grid(column=0, row=3, stick='w', pady=5, padx=13) 

    needstxt = Entry(window, width=30)  
    needstxt.grid(column=0, row=3, padx=420)
    
    
    def clear_window():
        progress_bar.destroy()
        lbl_min.destroy()
        lbl_need.destroy()
        lbl_start.destroy()
        needstxt.destroy()

    def compress():
        global progress_bar
        btn_compress.destroy()

        progress_bar = ttk.Progressbar(window, orient="horizontal", mode="determinate", maximum=100, value=0, length=450)
        progress_bar.grid(column=0, row=4, pady=65, padx = 120, sticky=W)
        progress_bar['value'] = 0
        window.update()


        stop = 0
        def error():
            stop = 1  
            clear_window()
            lbl_error = Label(text='Ошибка', fg = 'black', font=("Arial", 26), width=10)
            lbl_error.grid(padx=250, pady=120)

        try:
            need_size = int(needstxt.get())
        except ValueError:
            error()
        needstxt.config(state='disabled')

        if need_size > start_size or need_size < min_size:
            error()
        

        if stop == 0:
            for q in range(100, 1, -1):
                img = im.save(save_directory, quality = q, optimize = True)
                final_image = Image.open(save_directory)
                end_size = (os.stat(save_directory).st_size)//1000
                progress_bar['value'] = (100*((start_size-end_size)/(start_size-need_size)))
                window.update()
                if end_size < need_size or end_size == need_size:
                    clear_window()
                    lbl_dsksave = Label(text='Сжатое изображение сохранено на рабочий стол :)', fg = 'black', font=("Arial", 20), width=42)
                    lbl_dsksave.grid(column=2, pady=70, padx=10)
                    lbl_info = Label(text='Вес нового изображения - ' + str(end_size) + ' кб', fg = 'black', font=("Arial", 20), width=28)
                    lbl_info.grid(column=2, padx=20)
                    print('Используемый параметр качества -', q)
                    print('Конечный размер -', end_size, 'килобайт')
                    break

        

    btn_compress = Button(window, text="Сжать изображение", font=('Normal', 15), bg='#2f91ed', fg='White', height=3, width=20, command=compress)
    btn_compress.grid(column=0, row=5, padx=230, pady=35, sticky='w')


btn = Button(window, text="Выбрать изображение", font=('Normal', 19), bg='#2f91ed', fg='white', command = clicked)
btn.place(relx=0.5, rely=0.5, anchor=CENTER, height=100, width=350)

window.mainloop()

