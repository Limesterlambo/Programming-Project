Module Module1
    Dim t
    Dim current_line

    'to seperate your work out in the command prompt.
    Function Space(Number_of_Spaces)
        For i = 1 To Number_of_Spaces
            Console.Write(vbNewLine)
        Next
    End Function

    'General up to 5 column table maker.
    Function Table_Heading(Number_of_columnns As Integer, Heading1 As String, Heading2 As String, Heading3 As String, Heading4 As String, Heading5 As String)
        Dim n
        n = Number_of_columnns
        If n > 0 Then
            Console.Write(Heading1 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 1 Then
            Console.Write(Heading2 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 2 Then
            Console.Write(Heading3 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 3 Then
            Console.Write(Heading4 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 4 Then
            Console.WriteLine(Heading5 & vbTab & "|" & vbTab)
        Else
            Console.WriteLine("")
        End If
        'If the table devision is off, modify the number of times "-" string is duplicated.
        Console.WriteLine(StrDup((n * 17) + 7, "-"))
    End Function

    'time = distance/velocity rearranged from velocity = (distance/time), assuming SI units.
    Function Time_Taken(distance As Double, speed As Integer)
        t = (distance / speed)
    End Function

    'writing data to a text file on consecutive lines.
    Function Write_To_txt(file_path, Data)
        Dim file_writer = My.Computer.FileSystem.OpenTextFileWriter(file_path, True)

        file_writer.WriteLine(Data)

        file_writer.Close()
    End Function

    'Reading data from particular lines in a text file.
    Function Read_From_txt(file_path, line)
        Dim file_reader = My.Computer.FileSystem.OpenTextFileReader(file_path)
        Dim i = 0

        current_line = file_reader.ReadLine()

        While i < line
            current_line = file_reader.ReadLine()
            i = i + 1
        End While

    End Function

    'Average calculator.
    Function Average(Total, number_of_data_samples)
        Average = Total / number_of_data_samples
    End Function

    Sub Main()
        Dim material = {"Wood    ", "Steel   ", "Titanium"}
        Dim Average_speed = {323.5, 180.8, 100.3}
        Dim depth(2)
        Dim pre_average_time As Double 'the variable that will be used to calculate the average time.

        'q1 part c - fore user to input positive doubles. 
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & Trim(material(i)) & " you are testing on in meters:")
            depth(i) = Console.ReadLine
            ' Making sure the assighned value is a positive integer (not including 0 as the average speed infers it spends time in the medium).
            For i1 = 0 To 2
                If depth(i) <= 0 Then
                    ' If it isnt, asking the user to reinput a value
                    Console.WriteLine("please enter a possitive double for the depth of " & Trim(material(i)))
                    depth(i) = Console.ReadLine()
                End If
            Next
        Next

        Space(1)

        'q1 part a.
        '-part 1 - confirming users inputed depths with them.
        For i = 0 To 2
            Console.WriteLine("For distance in " & Trim(material(i)) & " you entered: " & depth(i) & " metres")
        Next

        Space(1)
        '-part 2 - forming a table for the time taken by the bullet to reach the target.

        '--Designing  the table columns.
        Table_Heading(2, "Material", "Time Taken (s)", "", "", "")

        '--Designing  the table rows.
        For i = 0 To 2
            Console.Write(material(i) & vbTab & "|" & vbTab)
            Time_Taken(depth(i), Average_speed(i))
            Console.WriteLine(Format(t, "0.000") & vbTab & vbTab & "|")

            'storing the data in a text file for later use.
            Write_To_txt("C:\Users\Nickq20\Documents\UNI\Programming\Course_Work\Q1-Code\Data File.txt", t)
        Next

        Space(1)

        'q1 part b.
        '-adding up pre stored times for each material so we can find the average.
        For i = 1 To 3
            Read_From_txt("C:\Users\Nickq20\Documents\UNI\Programming\Course_Work\Q1-Code\Data File.txt", i)
            pre_average_time = current_line + pre_average_time
        Next

        '-average calculation.
        Console.WriteLine("""The time taken by the bullet to reach the target on average"" = " & Format((Average(pre_average_time, 3)), "0.0000"))
    End Sub

End Module
