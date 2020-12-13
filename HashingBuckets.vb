Module Module1
    Dim StudentArray(5) As Integer
    Sub Main()
        Dim RecordKey, Counter As Integer
        For Counter = 0 To 5
            Console.Write("Enter record key : ")
            RecordKey = Console.ReadLine
            Call InsertHash(RecordKey)
        Next
        Console.Write("Enter record key to search for: ")
        RecordKey = Console.ReadLine
        Counter = SearchHash(RecordKey)
        If Counter = 0 Then
            Console.WriteLine("Record not found .")
        Else
            Console.WriteLine("Record found. ")
        End If
        Console.ReadKey()
        For Counter = 0 To 5
            Console.WriteLine(StudentArray(Counter))
        Next
        Console.ReadKey()
    End Sub
    Sub InsertHash(ByVal RecKey As Integer)
        Dim HashKey As Integer
        HashKey = Hash(RecKey, 5)
        While StudentArray(HashKey) > 0
            HashKey = HashKey + 1
            If HashKey > 5 Then
                HashKey = 0
            End If
        End While
        StudentArray(HashKey) = RecKey
    End Sub
    Function SearchHash(ByVal RecKey As Integer) As Integer
        Dim HashKey, TotalSearches As Integer
        HashKey = Hash(RecKey, 5)
        While StudentArray(HashKey) <> RecKey
            HashKey = HashKey + 1
            TotalSearches = TotalSearches + 1
            If HashKey > 5 Then
                HashKey = 0
                If TotalSearches > 5 Then
                    Return 0
                    Exit Function
                End If
            End If
        End While
        Return RecKey
    End Function
    Function Hash(ByVal KeyValue As Integer, ByVal MaxPosition As Integer) As Integer
        Dim Index As Integer
        Index = KeyValue Mod (MaxPosition + 1)
        Return Index
    End Function
End Module
