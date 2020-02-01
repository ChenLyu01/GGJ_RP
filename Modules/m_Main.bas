Attribute VB_Name = "m_Main"
Option Explicit
Public frmMain As frm_Main

Sub Main()
 
    
    this_FilePath.Graphics = App.Path & "\Graphics\Skins\"
    this_FilePath.CourseMap = App.Path & "\Map\"
    this_FilePath.Code = App.Path & "\Data\"
    this_FilePath.Story = App.Path & "\Story\"
    
    Set frmMain = New frm_Main
    Load frmMain
    Call Game_Initialize(frmMain.PicMain.hDC, this_FilePath)
    frmMain.Show
End Sub
