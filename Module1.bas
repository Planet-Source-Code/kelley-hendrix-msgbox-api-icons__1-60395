Attribute VB_Name = "Module1"
Public Enum MsgPicTypes
    Question = 32514&
    Exclamation = 32515&
    Critical = 32513&
    Information = 32516&
End Enum

Public Declare Function LoadStandardIcon Lib "user32" Alias _
    "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As _
    MsgPicTypes) As Long
    
Public Declare Function DrawIcon Lib "user32" (ByVal hDC _
    As Long, ByVal x As Long, ByVal y As Long, _
    ByVal hIcon As Long) As Long

