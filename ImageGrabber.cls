VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
'Downloads Image by URL and paste on worksheet to particular cell.
'Sets row height according to image resize height
'Returns error number
'
'Requirements:
'   - Microsoft WinHTTP Services, version 5.1 referense:
'     goto Tools->References menu to include
'
'==============================================================================
'EXAMPLE
'Puts image.jp to Cells(3,2) on Worksheets(1), resizes image height to 100
'
'Dim ImgGrab As ImageGrabber
'
'Set ImgGrab = New ImageGrabber
'
'ImgGrab.GetImage("https://www.site.com/image.jpg", Worksheets(1), 3, 2, 100)
'
'==============================================================================
'https://github.com/v-messer/
'==============================================================================

Dim WHTTP           As Object
Dim imageURL        As String
Dim imageLocalPath  As String


Private Sub Class_Initialize()

Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")

End Sub


Public Function GetImage( _
                            ByVal URL As String, _
                            ByRef ws As Worksheet, _
                            ByVal row As Long, _
                            ByVal column As Long, _
                            ByVal imageHeight As Integer _
                        ) As String
'Main image downloading method
'URL    - image URL
'ws     - worksheet ot paste
'row    - destionation cell row number
'column - destionation cell column number
Dim FileData() As Byte
Dim FileNum     As Long
Dim img         As Object

'destionation cell
ws.Cells(row, column).Select

WHTTP.Open "GET", URL, False
On Error Resume Next
WHTTP.Send

If Err.Number = 0 Then
    
    'if valid response
    If Not IsEmpty(WHTTP.ResponseBody) Then
        
        'download and saveimage
        FileData = WHTTP.ResponseBody
        Set WHTTP = Nothing
        FileNum = FreeFile
        imageLocalPath = "img.jpg"
        Open imageLocalPath For Binary Access Write As #FileNum
            Put #FileNum, 1, FileData
        Close #FileNum
        
        'check that file is image
        '----------------------------------------------------------------------
        'LoadPicture supports:
        '.bmp, .ico, .cur, .rle, .wmf, .emf, gif, .jpg.
        '.png is not supported
        'Returns Error 481 - Invalid Picture
        '----------------------------------------------------------------------
        LoadPicture imageLocalPath
        
        'if file is image
        If (Err.Number <> 481) Or (Right(URL, 4) = ".png") Then
            'insert image
            Set img = ActiveSheet.Pictures.Insert(imageLocalPath)
            img.Height = imageHeight
            'resize destionation cell row
            ActiveCell.RowHeight = imageHeight
        End If
        
        'delete temp image file
        SetAttr imageLocalPath, vbNormal
        Kill imageLocalPath
    
    End If 'Not IsEmptyWHTTP.ResponseBody)

End If 'Err.Number = 0

'return error number
GetImage = "Error " & Err.Number

End Function











