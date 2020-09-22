Attribute VB_Name = "ChatMod"
'******************************************************
'These variables are used in the load proccess
'******************************************************
'Start Variable(Used to increment the Select Case Statment)
Global Start_Sequence As Integer
'File Variable(Contents of the file is put in this)
Global Start_FileOpen As String
'Size Variable(Gets the length of the string)
Global Start_FileSize As Integer
'FilePath Variable(Gets the compleat file path with the carrige return chars taken out)
Global Start_FilePath As String
'******************************************************
'Variables used during operation proccess
'******************************************************
'Decision Variable(Used to tell whether the path is needed or not)
Global Op_PathNeeded As Integer
'Path Variable(Holds the path of where the chat rooms are to be)
Global Op_ChatPath As String
'******************************************************
'Chat Room Names(Used to join the rooms)
'******************************************************
Global Room_General As String
Global Room_Hacking As String
Global Room_Work As String
Global Room_Life As String
Global Room_Lobby As String
'******************************************************
'******************************************************
'Used to open desired room
Global set_room As String
'Used to store current room
Global Current_Room As String
'Used to store Temp files
Global File_Temp As String
'Used to keep the users name
Global NameOfPerson As String
'Used to hold users text file
Global Set_Users As String
'Used to hold the users names
Global Users_Names As String
'Used to store the path to the all users.txt
Global All_Users As String
'******************************************************
'Chat Room users(Used to store them)
'******************************************************
Global Users_General As String
Global Users_Hacking As String
Global Users_Work As String
Global Users_Life As String
Global Users_Lobby As String
'******************************************************
'******************************************************
'Used to store tempory users to see if the user file contains anythink
Global Temp_User As String
'Used to cout up for room clean
Global Counters As Integer
'Used to check whether the rooms ar empty
Global Temp_RoomCheck As String
'Used to store what the user types in the textbox
Global User_Message As String
'Used to store the file-path at start up
Global Temp_Filepath As String
'Used to stop some error handlers
Global Stop_Handle As Integer
'Hummmmmmmm
Global Temp_F As String
'Used to stop auto scroll
Global Auto_Scroll As Integer
'Used to change users name
Global Change_Name As String
'used to store the users firts name before changing it
Global first_name As String
