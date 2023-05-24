VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If
' ----------------------------------------------------------------
' Дата: 05.04.2023 21:08
' Автор: akolodka
' ----------------------------------------------------------------

Const SHOW_MINCOUNT = 10
Const DEFAULT_CAPTION = "Кеширование данных..."
Const FORMAT_PERCENT = "0 %"
Const SLEEP_INTERVAL = 400

Private Type TStatusBar
    
    caption As String
    
    count As Long
    step As Double
    
    widthProgress As Double
    percent As Double
    
    isLoaded As Boolean
    
End Type

Private this As TStatusBar

Private Sub Class_Initialize()
    
    this.caption = DEFAULT_CAPTION
    this.isLoaded = False
    
End Sub
' ----------------------------------------------------------------
' Дата: 05.04.2023 21:14
' ----------------------------------------------------------------
Public Sub Initialize( _
    totalCount As Long, _
    Optional captionText As String _
    )

    Dim isShowStatusBar As Boolean
    isShowStatusBar = totalCount > SHOW_MINCOUNT
    
    If Not isShowStatusBar Then _
        Exit Sub
        
    this.count = totalCount
    this.step = frmStatusBar.frameBackground.Width / totalCount
    
    this.percent = vbEmpty
    this.widthProgress = vbEmpty
        
    If captionText <> vbNullString Then _
        this.caption = captionText
        
    frmStatusBar.labelBlack.Left = frmStatusBar.labelWhite.Left
    frmStatusBar.caption = this.caption
    
    frmStatusBar.Show False
    this.isLoaded = True
    
End Sub
' ----------------------------------------------------------------
' Дата: 05.04.2023 21:25
' ----------------------------------------------------------------
Public Sub ShowProgress()

    this.widthProgress = this.widthProgress + this.step
    frmStatusBar.frameProgress.Width = this.widthProgress
     
    If this.widthProgress <= this.percent Then Exit Sub
        
    Dim progressPercent As String
    progressPercent = Format(this.percent / frmStatusBar.frameBackground.Width, FORMAT_PERCENT)
    
    frmStatusBar.labelWhite.caption = progressPercent
    frmStatusBar.labelBlack.caption = progressPercent
    
    this.percent = this.percent + this.step
    frmStatusBar.Repaint
    
End Sub
' ----------------------------------------------------------------
' Дата: 05.04.2023 22:01
' ----------------------------------------------------------------
Public Sub Terminate()
    
    Sleep SLEEP_INTERVAL
    
    If this.isLoaded Then _
        VBA.Unload frmStatusBar
        
End Sub
