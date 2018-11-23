﻿Module Module1
    Function Table_Headings() 'no inputs required'
        'making the headings of the table'
        Console.WriteLine("Material /" & vbTab & "Depth (meters) /" & vbTab & "Average Speed (m/s) /" & vbTab & "Time Spent in the medium (s)")
        Console.WriteLine("---------------------------------------------------------------------------------------------")
    End Function

    Sub Main()
        'Pre given values placed in acording arrays'
        Dim average_speeds() As Double = {323.5, 180.8, 100.3}
        Dim material() As String = {"Wood    ", "Steel   ", "Titanium"}
        'A variable for the user to input a depth for ecah material'
        Dim Depth(2) As Double
        'a usfull index for out for loops to work'
        Dim i, i1

        'asking the user to input a depth and then assighning that value to a possition in an array'
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & Trim(material(i)) & " you are testing on in meters")
            Depth(i) = Console.ReadLine
            'making sure the assighned value is a positive integer (not including 0 as the average speed infers it spends time in the medium)'
            For i1 = 0 To 2
                If Depth(i) <= 0 Then
                    'If it isnt, asking the user to input a value in again'
                    Console.WriteLine("please enter a possitive double for the depth of " & Trim(material(i)))
                    Depth(i) = Console.ReadLine()
                End If

            Next
        Next

        'outputing the table of results'
        Table_Headings()

        'outputing the rows'
        For i = 0 To 2
            Console.WriteLine(material(i) & " /" & vbTab & Space(3) & Format(Depth(i), "00.000") & Space(6) & "/" & Space(10) & Format(average_speeds(i), "0.000") & Space(11) & "/" & Space(11) & Format(Depth(i) / average_speeds(i), "0.000"))
        Next

    End Sub

End Module
