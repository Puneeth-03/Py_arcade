import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *
def number_guess():
    import Number_guessing
    messagebox.showinfo("Game Location",'You have to play in IDLE')
    print("---------------GUESS THE NUMBER GAME----------------")
    Number_guessing.start_game()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')
def dots_boxes():
    import dotsandboxes
    messagebox.showinfo("Game Location",'You have to play in IDLE')
    print("---------------DOTS AND BOXES GAME----------------")
    dotsandboxes.instructions()
    dotsandboxes.initialize()
    dotsandboxes.start_game()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')
def hangman():
    import hangman
    hangman.start_game(root)
def madlibs():
    import mad_libs
    messagebox.showinfo("Game Location",'You have to play in IDLE')
    print("---------------MAD LIBS GAME----------------")
    mad_libs.start_game()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')
def Guess_the_colour():
    import Guessthecolour
    Guessthecolour.startGame(event)
    Guessthecolour.nextcolour()
    Guessthecolour.countdown()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')
    
def Tic_tac_toe():
    import TicTacToe
    TicTacToe.winner(b,1)
    TicTacToe.get_text(i,j,gb,l1,l2)
    TicTacToe.isfree(i,j)
    TicTacToe.isfull()
    TicTacToe.gameboard_pl(game_board,l1,l2)
    TicTacToe.pc()
    TicTacToe.get_text_pc(i,j,gb,l1,l2)
    TicTacToe.gameboard_pc(game_board,l1,l2)
    TicTacToe.withpc(game_board)
    TicTacToe.withplayer(game_board)
    '''reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')'''
def escape_room():
    import escape_room
    messagebox.showinfo("Game Location",'You have to play in IDLE')
    print("---------------ESCAPE ROOOM GAME----------------")
    escape_room.start_game()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        Message(root,text='Click on EXIT')
def rock_paper_scissors():
    import Rock_paper_scissor
    messagebox.showinfo("Game Location",'You have to play in IDLE')
    print("---------------ROCK, PAPER AND SCISSORS GAME----------------")
    Rock_paper_scissor.start_game()
    reply=messagebox.askquestion(title,text)
    if reply=='yes':
        messagebox.showinfo("Continue",'Click on a GameIcon')
    else:
        messagebox.showinfo("EXIT",'Click on EXIT')
def game_result():
    import Result
    
root=Tk()
root.geometry('1350x1000')
root.title("PY-ARCADE")
bg=PhotoImage(file="background4.png")
my_label=Label(root,image=bg)
my_label.place(x=0,y=0,relwidth=1,relheight=1)
label1=tkinter.Label(root,text="WELCOME TO PY-ARCADE!!!\nCLICK ON A GAMEICON TO PLAY!!\nHAVE FUN !!!",font=('Comic Sans MS',15),fg='black',bg='#ffccff').grid(row=0,column=2)

photo1=PhotoImage(file =r"Guess_my_number.png")
photo2=PhotoImage(file =r"Dots_and_boxes_game.png")
photo3=PhotoImage(file =r"hangman.png")
photo4=PhotoImage(file =r"mad-libs.png")
photo5=PhotoImage(file =r"Guess the colour(1).png")
photo6=PhotoImage(file =r"tic tac toe(1).png")
photo7=PhotoImage(file =r"escape_room.png")
photo8=PhotoImage(file =r"rock-paper-scissors.png")

Button(root,text="Guess my number",image=photo1,command=number_guess).grid(row=5,column=0)
Button(root,text="Dots and boxes",image=photo2,command=dots_boxes).grid(row=5,column=1)
Button(root,text="Hangman",image=photo3,command=hangman).grid(row=6,column=0)
Button(root,text="Mad Libs",image=photo4,command=madlibs).grid(row=6,column=1)
Button(root,text="Guess the colour",image=photo5,command=Guess_the_colour).grid(row=5,column=2)
Button(root,text="Tic tac toe",image=photo6,command=Tic_tac_toe).grid(row=5,column=3)
Button(root,text="Escape room",image=photo7,command=escape_room).grid(row=6,column=2)
Button(root,text="Rock paper scissors",image=photo8,command=rock_paper_scissors).grid(row=6,column=3)

b1=tkinter.Button(root,text='EXIT',height=2,width=20,fg='black',bg='#ffccff',font=('Comic Sans MS',15,'bold'),command=root.destroy)
b1.place(x=440,y=525)
b2=tkinter.Button(root,text='RESULTS',height=2,width=20,fg='black',bg='#ffccff',font=('Comic Sans MS',15,'bold'),command=game_result)
b2.place(x=700,y=525)

title= 'Game Continuation Confirmation'
text = 'Do you want to play more?'

root.mainloop()
