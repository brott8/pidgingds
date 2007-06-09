Public Class clsCRC32

    ' This is v2 of the VB CRC32 algorithm provided by Paul
    ' (wpsjr1@succeed.net) - much quicker than the nasty
    ' original version I posted.  Excellent work!

    Private crc32Table() As Integer
    Private Const BUFFER_SIZE As Integer = 1024

    Public Function GetCrc32(ByRef str As String) As Integer
        Dim encoding As New System.Text.ASCIIEncoding()
        Dim bytearray As Byte() = encoding.GetBytes(Str)

        Dim crc32Result As Integer
        crc32Result = &HFFFFFFFF

        'Dim count As Integer = stream.Read(Buffer, 0, readSize)
        Dim i As Integer
        Dim iLookup As Integer
        Dim tot As Integer = 0
        'Do While (count > 0)
        For i = 0 To bytearray.Length - 1
            iLookup = (crc32Result And &HFF) Xor bytearray(i)
            crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And &HFFFFFF ' nasty shr 8 with vb :/
            crc32Result = crc32Result Xor crc32Table(iLookup)
        Next i
        'count = stream.Read(Buffer, 0, readSize)
        'Loop

        GetCrc32 = Not (crc32Result)

    End Function

    Public Sub New()

        ' This is the official polynomial used by CRC32 in PKZip.
        ' Often the polynomial is shown reversed (04C11DB7).
        Dim dwPolynomial As Integer = &HEDB88320
        Dim i As Integer, j As Integer

        ReDim crc32Table(256)
        Dim dwCrc As Integer

        For i = 0 To 255
            dwCrc = i
            For j = 8 To 1 Step -1
                If (dwCrc And 1) Then
                    dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    dwCrc = dwCrc Xor dwPolynomial
                Else
                    dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                End If
            Next j
            crc32Table(i) = dwCrc
        Next i
    End Sub

End Class
