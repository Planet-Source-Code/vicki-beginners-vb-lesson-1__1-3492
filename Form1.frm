VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   3555
   ClientTop       =   2715
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

'Always start ALL routines by declaring the variables you will be using!
    Dim the_secret_number As Long
    Dim counter As Long
    Dim users_guess_string As String
    Dim users_guess_number As Long

'First we'll randomize the timer so that we get a different
'random number every time we play.

    Randomize Timer

'Next we'll generate our random number between 1 and 100

    the_secret_number = Int((100 * Rnd) + 1)

'Now we'll set up a For-Next loop which will give the user 10 tries at
'guessing the secret number.

For counter = 1 To 10
    
    'Have the user take a guess.
    users_guess_string = InputBox("Take a guess between 1 and 100")
    'Let's convert the users guess from a string to a number.
    users_guess_number = Val(users_guess_string)
    'Let's compare his guess with the secret number.
    If users_guess_number < the_secret_number Then
        MsgBox (users_guess_number & " is too low")
    End If
    
    If users_guess_number > the_secret_number Then
        MsgBox (users_guess_number & " is too high")
    End If
    
    If users_guess_number = the_secret_number Then
        MsgBox (users_guess_number & " was the number! Congratulations!")
        Exit Sub 'We have to exit the sub here, or we'll continue with the For-Next loop.
    End If
    
Next counter

End Sub


