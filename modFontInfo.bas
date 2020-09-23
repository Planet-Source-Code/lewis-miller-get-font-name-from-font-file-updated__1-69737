Attribute VB_Name = "modFontInfo"
Option Explicit

'******************************************************************
'thanks to Philip Patrick and his c++ article on the Code Project
'http://www.codeproject.com/KB/GDI/fontnamefromfile.aspx
'VB code by Lewis Miller 12/04/07 dethbomb@hotmail.com
'******************************************************************

'12/11/07 BugFix: Soorya has noted a problem with unicode font names. His proposed fix
'                 has been added. Thanks to soorya for the bug find.

'Remarks:
'Font files store all there information in motorola (or Big-Endian) format
'which is incompatible with vb, so we must use memory swapping tricks
'to retrieve the values we want to use from font files. You cannot access/read a variable
'that has been loaded from the font file unless you first swap it to intel (or Little-Endian)
'format.... doing so will cause havoc in your program :)

'api declarations
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal Length As Long)
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'font structures (taken from philips article)
Public Type OFFSET_TABLE    'len = 12
    uMajorVersion As Integer
    uMinorVersion As Integer
    uNumOfTables  As Integer
    uSearchRange  As Integer
    uEntrySelector As Integer
    uRangeShift   As Integer
End Type

Public Type TABLE_DIRECTORY    'len = 16
    szTag         As String * 4    'table name
    uCheckSum     As Long    'Check sum
    uOffset       As Long    'Offset from beginning of file
    uLength       As Long    'length of the table in bytes
End Type

Public Type NAME_TABLE_HEADER    'len = 6
    uFSelector    As Integer    'format selector. Always 0
    uNRCount      As Integer    'Name Records count
    uStorageOffset As Integer    'Offset for strings storage, from start of the table
End Type

Public Type NAME_RECORD    'len = 12
    uPlatformID   As Integer    '
    uEncodingID   As Integer    '
    uLanguageID   As Integer    '
    uNameID       As Integer    '
    uStringLength As Integer    '
    uStringOffset As Integer    ' //from start of storage area
End Type


'************************************************************
'Helper Functions
'***********************************************************

'convert a big-endian Long to a little-endian Long
Sub SwapLong(LongVal As Long)
    LongVal = ntohl(LongVal)
End Sub

'convert a big-endian Integer (short) to a little-endian Integer (short)
Sub SwapInt(IntVal As Integer)
    IntVal = ntohs(IntVal)
End Sub


'************************************************************
'retrieves the font name from a font file
'the file must be a true type font 1.0
'***********************************************************

Function GetFontName(ByVal FontPath As String) As String

    Dim TblDir      As TABLE_DIRECTORY 'table directory
    Dim OffSetTbl   As OFFSET_TABLE    'table information
    Dim NameTblHdr  As NAME_TABLE_HEADER 'name table info
    Dim NameRecord  As NAME_RECORD       'info table
    Dim FileNum     As Integer
    Dim lPosition   As Long
    Dim sFontTest   As String
    Dim X           As Long
    Dim I           As Long
    
    'make sure font file exists
    If Dir$(FontPath) = vbNullString Then
       Exit Function
    End If
    
    'open the file
    On Error GoTo Finished
    FileNum = FreeFile
    Open FontPath For Binary As FileNum

    'read the first main table header
    Get #FileNum, , OffSetTbl

    'check major and minor versions for 1.0
    With OffSetTbl
        SwapInt .uMajorVersion
        SwapInt .uMinorVersion
        If .uMajorVersion <> 1 Or .uMinorVersion <> 0 Then
            'MsgBox "Invalid font file version, cannot read font name!", vbCritical
            GoTo Finished
        End If
        SwapInt .uNumOfTables
    End With

    If OffSetTbl.uNumOfTables > 0 Then
        For X = 0 To OffSetTbl.uNumOfTables - 1
            Get #FileNum, , TblDir
            If StrComp(TblDir.szTag, "name", vbTextCompare) = 0 Then
                'we have found the name table hdr, now we get the length and offset of name record
                With TblDir
                    SwapLong .uLength
                    SwapLong .uOffset
                    If .uOffset Then
                        Get #FileNum, .uOffset + 1, NameTblHdr
                        SwapInt NameTblHdr.uNRCount
                        SwapInt NameTblHdr.uStorageOffset

                        For I = 0 To NameTblHdr.uNRCount - 1
                            Get #FileNum, , NameRecord
                            SwapInt NameRecord.uNameID
                            '1 specifies font name, this could be modified to get other info
                            If NameRecord.uNameID = 1 Then
                                SwapInt NameRecord.uStringLength
                                SwapInt NameRecord.uStringOffset
                                lPosition = Loc(FileNum)    'save current file position

                                If NameRecord.uStringLength Then
                                    sFontTest = Space$(NameRecord.uStringLength)
                                    Get #FileNum, .uOffset + NameRecord.uStringOffset + NameTblHdr.uStorageOffset + 1, sFontTest
                                    If Len(sFontTest) Then
                                        GoTo Finished    'all done
                                    End If
                                End If

                                'string was empty so , search more
                                Seek #FileNum, lPosition

                            End If
                        Next I
                    End If
                End With
            End If
        Next X
    End If


Finished:
    Close #FileNum
    
    'note: some fonts are returned in unicode (double byte) format
    '      so we must remove the null characters - thanks to soorya
    GetFontName = Replace$(sFontTest, vbNullChar, "")

End Function


