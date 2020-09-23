VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Secure Delete"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete File"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' secure file delete by the laughing man (AKA Travis)
'
'   This is a program that will securely erase a file. This program was coded so
' that it can be copied into a module without needing to make modification. this
' method of mine works by randomly generating a random character then overwriting
' every character in the file with it, Then on the next pass over use a new character.
' This program will let the user pick how many passes to make on the file as well. The code is commented but please let me know if you have further questions. Also this is my first upload, please let me know what you think and Enjoy!!!
'
'

Dim target As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function SecureDeleteFile(Path As String, NumOfPasses As Integer)
    Dim i As Long 'Loop Control
    Dim Data As String 'Data to OverWrite the File with
    Dim Lenght As Long 'Var to hold Lenght OF File in to prevent overflows in the loop
    Dim upperbound As Long ' used in random char generator. this is the high value
    Dim lowerBound As Long ' used in random char generator. this is the low value
    Randomize ' Intalize Random
    upperbound = 255 ' highest value allowed for a random char (ascii is 0 - 255)
    lowerBound = 0 'lowest value allowed for an ascii value
    
    Data = Chr(CInt(Int((upperbound - lowerBound + 1) * Rnd() + lowerBound))) ' This
                            'generates a ascii charcter out of a number 0-255 that
                            'is randomly picked by the machine

For j = 1 To NumOfPasses 'loop that is used to control the number of overwrites on the file
    Open Path For Binary Access Write As #1 ' open the file
    Lenght = LOF(1) ' find the lenght of the file and place its value into the var "Lenght"
    i = 0 ' zero out the loop counter, note that the For loop I used "J"
Test: ' a call to location for my loop to use
        If i <> Lenght Then ' i is increased with everyloop untill it meets Lenght
            Put #1, , Data ' place the random charcter in to the file
            i = i + 1 ' increase the i control, this is so that it will overwrite every byte in the file
            GoTo Test: ' duh!
        Else
        End If
    Close #1 ' close the file to ensure that it is written to the disk and to prevent any problems
    Data = Chr(CInt(Int((upperbound - lowerBound + 1) * Rnd() + lowerBound))) ' generate a new char for the next loop
Next j ' proceed the next pass of the file or quit
Sleep 500 ' this waits so the computer hardware can catch up a little. this was the main problem at first.
            ' I was able to recover files on a single pass that were only half destroyed.
Kill Path ' delete the file after the two loops have destroyed the contents of the file.
MsgBox "Done!" ' a little user interaction never hurts
End Function

Private Sub Command2_Click()
CommonDialog1.ShowOpen ' opens a dialog to allow user to select there file
Text1.Text = CommonDialog1.FileName ' places the name of the selected file in the textbox
End Sub

Private Sub Command1_Click()
Dim NumbofPass As Integer
target = Text1.Text
NumbofPass = Text2.Text
If Text1.Text = "" Then
    MsgBox "Please Select a File" ' error if user tries to run without a file selected.
Else
    If NumbofPass = 0 Then ' if user fails to input a number of passes to make.
        NumbofPass = 10   ' I choice ten because at 9 the file will be wipe to the point
                            ' that it becomes difficult to recover on a physical level as well
    Else
    End If
    SecureDeleteFile target, NumbofPass 'call the function
End If
End Sub
