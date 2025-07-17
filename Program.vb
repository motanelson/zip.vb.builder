Imports System
Imports System.IO
Imports System.Text


Module ZipCreator

    ' === Tabela CRC32 ===
    Dim crc32Table(255) As UInteger

    Sub InitCRC32Table()
        For i As Integer = 0 To 255
            Dim c As UInteger = CUInt(i)
            For j As Integer = 0 To 7
                If (c And 1) <> 0 Then
                    c = &HEDB88320UI Xor (c >> 1)
                Else
                    c >>= 1
                End If
            Next
            crc32Table(i) = c
        Next
    End Sub

    Function CalcCRC32(data As Byte()) As UInteger
        Dim crc As UInteger = &HFFFFFFFFUI
        For Each b In data
            crc = (crc >> 8) Xor crc32Table((crc Xor b) And &HFF)
        Next
        Return Not crc
    End Function

    Sub WriteU16(stream As Stream, value As UShort)
    stream.WriteByte(CByte(value And &HFF))
    stream.WriteByte(CByte((value >> 8) And &HFF))
End Sub

Sub WriteU32(stream As Stream, value As UInteger)
    stream.WriteByte(CByte(value And &HFF))
    stream.WriteByte(CByte((value >> 8) And &HFF))
    stream.WriteByte(CByte((value >> 16) And &HFF))
    stream.WriteByte(CByte((value >> 24) And &HFF))
End Sub
    Sub Main()
        InitCRC32Table()
        Console.ForegroundColor = ConsoleColor.Black
        Console.BackgroundColor = ConsoleColor.DarkYellow
        Console.Write("Ficheiros a empacotar (separados por espaço): ")
        Dim input As String = Console.ReadLine()
        Dim files() As String = input.Split(" "c)
        Dim zipName As String = "output.zip"

        Using zip As New FileStream(zipName, FileMode.Create)
            Dim offsets As New List(Of UInteger)
            Dim centralDirs As New List(Of Byte())
            Dim encoding As Encoding = Encoding.ASCII

            For Each file In files
                If Not  System.IO.File.Exists(file) Then
                    Console.WriteLine("Ignorado: " & file)
                    Continue For
                End If

                Dim data As Byte() =  System.IO.File.ReadAllBytes(file)
                Dim crc As UInteger = CalcCRC32(data)
                Dim nameBytes As Byte() = encoding.GetBytes(Path.GetFileName(file))
                Dim offset As UInteger = CUInt(zip.Position)
                offsets.Add(offset)

                ' --- Local File Header ---
                zip.Write(New Byte() {&H50, &H4B, &H3, &H4}, 0, 4) ' "PK\x03\x04"
                WriteU16(zip, 20) ' Version needed
                WriteU16(zip, 0)  ' Flags
                WriteU16(zip, 0)  ' Compression = store
                WriteU16(zip, 0)  ' Time
                WriteU16(zip, 0)  ' Date
                WriteU32(zip, crc)
                WriteU32(zip, CUInt(data.Length))
                WriteU32(zip, CUInt(data.Length))
                WriteU16(zip, CShort(nameBytes.Length))
                WriteU16(zip, 0)  ' Extra field
                zip.Write(nameBytes, 0, nameBytes.Length)
                zip.Write(data, 0, data.Length)

                ' --- Central Directory Header (buffered) ---
                Using ms As New MemoryStream()
                    ms.Write(New Byte() {&H50, &H4B, &H1, &H2}, 0, 4) ' "PK\x01\x02"
                    WriteU16(ms, &H0314)  ' Version made by (DOS + version 20)
                    WriteU16(ms, 20)      ' Version needed
                    WriteU16(ms, 0)       ' Flags
                    WriteU16(ms, 0)       ' Compression
                    WriteU16(ms, 0)       ' Time
                    WriteU16(ms, 0)       ' Date
                    WriteU32(ms, crc)
                    WriteU32(ms, CUInt(data.Length))
                    WriteU32(ms, CUInt(data.Length))
                    WriteU16(ms, CShort(nameBytes.Length))
                    WriteU16(ms, 0)       ' Extra
                    WriteU16(ms, 0)       ' Comment
                    WriteU16(ms, 0)       ' Disk number
                    WriteU16(ms, 0)       ' Internal attr
                    WriteU32(ms, 0)       ' External attr
                    WriteU32(ms, offset)  ' Offset of local header
                    ms.Write(nameBytes, 0, nameBytes.Length)
                    centralDirs.Add(ms.ToArray())
                End Using
            Next

            ' --- Central Directory ---
            Dim cdStart As UInteger = CUInt(zip.Position)
            For Each entry In centralDirs
                zip.Write(entry, 0, entry.Length)
            Next
            Dim cdSize As UInteger = CUInt(zip.Position - cdStart)

            ' --- End of Central Directory ---
            zip.Write(New Byte() {&H50, &H4B, &H5, &H6}, 0, 4) ' "PK\x05\x06"
            WriteU16(zip, 0) ' Disk number
            WriteU16(zip, 0) ' Disk with central dir
            WriteU16(zip, CShort(centralDirs.Count))
            WriteU16(zip, CShort(centralDirs.Count))
            WriteU32(zip, cdSize)
            WriteU32(zip, cdStart)
            WriteU16(zip, 0) ' Comment length
        End Using

        Console.WriteLine("ZIP criado: " & zipName)
    End Sub

End Module
