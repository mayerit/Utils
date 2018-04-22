Attribute VB_Name = "Utils"
Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

'CREATE TABLE [dbo].[employee](
'    [id] [varchar](36) NULL,
'    [firstname] [varchar](36) NULL,
'    [lastname] [varchar](36) NULL,
'    [designation] [varchar](36) NULL,
'    [intvalue] [int] NULL,
'    [decvalue] [decimal](19, 4) NULL,
'    [datevalue] [datetime] NULL
') ON [PRIMARY]




Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Public Function GetGUID() As String
'(c) 2000 Gus Molina

Dim udtGUID As GUID
'******************************************
Dim strDATA1 As String, strDATA2 As String, strDATA3 As String, strDATA4 As String, strDATA5 As String

If (CoCreateGuid(udtGUID) = 0) Then

GetGUID = _
String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & "-" & _
String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & "-" & _
String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & "-" & _
IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))

GetGUID = Left(GetGUID, 19) & Mid(GetGUID, 20, 4) & "-" & Mid(GetGUID, 24, 12)
GetGUID = LCase$(GetGUID)

End If

End Function

