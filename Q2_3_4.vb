Module Module1

    Sub Main()
        Dim input, min, max, sum, count1, count2, countfor, numbers(10), totals(9) As Integer
        Dim avg As Double
        count1 = 0
        count2 = 0
        max = 0
        min = 10

        Console.WriteLine("Enter your chosen 10 integers:")

        For countfor = 1 To 10
            input = Console.ReadLine()
            numbers(count1) = input
            If input > max Then
                max = input
            End If
            If input < min Then
                min = input
            End If
            count1 = count1 + 1
        Next

        count1 = 0

        While count2 < 9
            If numbers(count1) = count2 + 1 Then
                totals(count2) = totals(count2) + 1
                count1 = count1 + 1
            Else
                count2 = count2 + 1
            End If
        End While

        count1 = 0
        count2 = 0

        For countfor = 1 To 9
            Console.WriteLine("The total number of " & (count1 + 1) & "s was " & totals(count1))
            count1 = count1 + 1
        Next

        For countfor = 1 To 10
            sum = sum + numbers(count2)
            count2 = count2 + 1
        Next

        avg = sum / 10

        Console.WriteLine("The maximum value was " & max & vbCrLf & "The minimum value was " & min)
        Console.WriteLine("The average value was " & avg)
    End Sub

End Module