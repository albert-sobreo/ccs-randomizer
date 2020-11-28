from tkinter import *
import pandas as pd
import random
import csv

window = Tk()
window.configure(background="black")
window.title("CCS Randomizer")

w_h = 600
w_w = 800

small = "D:\Albert's Files\Programming\Randomizer\CCSRandomizer-2-small.png"
large = "D:\Albert's Files\Programming\Randomizer\CCSRandomizer-2.png"
medium = "D:\Albert's Files\Programming\Randomizer\CCSRandomizer-2(1).png"

font_size = "25"

canvas = Canvas(window, width=w_w, height=w_h, bg='black', highlightthickness=0)
canvas.pack()

image = PhotoImage(file=small)

bg_img = image
bg_label = canvas.create_image((0,0), image=bg_img, anchor=NW)

label_msg=canvas.create_text((w_w/2, w_h/2), text="", font="Bahnschrift " + font_size + " bold", fill="#ffffff")
label_last_winner=canvas.create_text((0,0), text="Last Winner: ", font="Bahnschrift 12 bold", fill="#ffffff", anchor=NW)

excel_file = "list of students 2020.xlsx"

def change_file(event):
    global excel_file
    if excel_file == "list of students 2020.xlsx":
        excel_file = "list of amazing students2020.xlsx"

    elif excel_file == "list of amazing students2020.xlsx":
        excel_file = "special.xlsx"

    else:
        excel_file = "list of students 2020.xlsx"

def main(event):
    dg = pd.read_excel("list of students 2020.xlsx", sheet_name=0)
    df = pd.read_excel(excel_file, sheet_name=0)
    mylist2 = dg['Students'].tolist()
    mylist = df['Students'].tolist()

    winner = ""

    with open('winners.csv', newline='') as f:
        reader = csv.reader(f)
        data = list(reader)

    ctr = 0
    while [winner] in data:
        if ctr > 50:
            return
        random_integer = random.randint(0,len(mylist)-1)
        winner = mylist[random_integer]
        ctr += 1

    delta = 20
    delay = 0

    for person in mylist2:
        update_text = lambda person=person: canvas.itemconfig(label_msg, text=person)
        canvas.after(delay, update_text)

        delay += delta

    with open('winners.csv', 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([winner])
        
    update_text = lambda winner=winner: canvas.itemconfig(label_msg, text=winner)
    canvas.after(delay, update_text)
    update_last_winner = lambda last_winner=data[-1][0]: canvas.itemconfig(label_last_winner, text="Last Winner: " + last_winner)
    canvas.after(delay, update_last_winner)

def delete(event):
    canvas.itemconfig(label_msg, text="")

    with open('winners.csv', newline='') as f:
        reader = csv.reader(f)
        data = list(reader)

    update_text = lambda s=data[-1][0]: canvas.itemconfig(label_last_winner, text="Last Winner: " + s)
    canvas.after(0, update_text)

canvas.bind("<Button-1>", main)
canvas.bind("<Button-2>", change_file)
canvas.bind("<Button-3>", delete)

window.mainloop()