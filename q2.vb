Module Module1
    Dim t

    '  Function To seperate work out in the console.
    Function Newlines(Number_of_newlines As Integer)
        For i = 1 To Number_of_newlines
            Console.Write(vbNewLine)
        Next
    End Function

    '  Fuction for generating table column headings for up to 5 columns
    Function Table_Heading(Number_of_columnns As Integer, Heading0 As String, Heading1 As String, Heading2 As String, Heading3 As String, Heading4 As String)
        Dim Headings(4) As String
        Headings(0) = Heading0
        Headings(1) = Heading1
        Headings(2) = Heading2
        Headings(3) = Heading3
        Headings(4) = Heading4

        ' If the table division is off, modify the number of times "-" string is repeated.
        Console.WriteLine(StrDup((Number_of_columnns * 17) + 7, "="))
        For I = 0 To 4
            If Number_of_columnns > I Then
                Console.Write(Space(2) & Headings(I) & vbTab & "|")
            Else
                Console.Write("")
            End If
        Next
        Console.WriteLine(vbNewLine & StrDup((Number_of_columnns * 17) + 7, "="))
    End Function

    ' Time = distance/velocity (rearranged from velocity = (distance/time), assuming SI units.)
    Function Calculate_Time_Taken(distance, speed) As Double
        Calculate_Time_Taken = (distance / speed)
    End Function

    ' Average calculator.
    Function Average(Total, number_of_data_samples) As Double
        Average = Total / number_of_data_samples
    End Function

    Sub Main()
        Dim material = {"Wood    ", "Steel   ", "Titanium"}
        Dim Average_speed = {323.5, 180.8, 100.3}
        Dim Average_Time_taken(6) As Double

        Dim depth(2)
        Dim j
        Dim Total_time_Combination As Double
        Dim Total_Time

        'Array to contain, sorted from smallest to largest, average time taken(s) of materials in each combination.
        Dim Sorted_Average_Time_Taken(6) As Double
        ' Creating an array containing all possible combination titles that will be needed for the table.
        Dim Combination_Titles() As String = {"Blanks and brown gunpowder" & Space(15), "Lead and brown gunpowder" & Space(20), "Armor-piercing and brown gunpowder" & Space(10),
            "Blanks and sulphur-free gunpowder" & Space(8), "Lead and sulphur-free gunpowder" & Space(10), "Armor-piercing and sulphur-free gunpowder", "No modifications" & Space(25)}

        ' Creating three arrays, one for each material, to store the modified setup's average speeds. This array will be be used to calculate the times taken.
        Dim Mod_Average_Speed_Wood(2, 1), Mod_Average_Speed_Steel(2, 1), Mod_Average_Speed_Titanium(2, 1) As Double

        'Creating arrays for time taken in each material.
        Dim Time_Taken_Wood(6), Time_Taken_Steel(6), Time_Taken_Titanium(6) As Double

        Dim l As Integer

        ' Array Bullet() elements correspond To 0 blank, 1 LeadBullets, 2 Armour-piercing Bullets.
        ' Array Gunpowder() elements correspond to 0 Brown Gunpowder, 1 Sulphur-free Gunowder.
        Dim Wood_Bullets() As Integer = {-3, 20, -30}
        Dim Wood_Gunpowder() As Integer = {35, -10}
        Dim Steel_Bullets() As Integer = {-5, 12, 5}
        Dim Steel_Gunpowder() As Integer = {35, -10}
        Dim Titanium_Bullets() As Integer = {-12, 5, 30}
        Dim Titanium_Gunpowder() As Integer = {35, -10}


        ' Q1 part (c) - forcing the user to input positive doubles. 
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
            depth(i) = Console.ReadLine

            ' Loop Making sure the assigned value is a number.
            Int32.TryParse(depth(i), j)
            While j = False
                Console.WriteLine("ERROR: depth must be a numerical! " &
                    "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
                Int32.TryParse(depth(i), j)
            End While
            depth(i) = CDbl(depth(i))

            ' Making sure the value is positive. (Not including 0 as the average speed infers it spends time in the medium).
            While depth(i) <= 0
                ' If it isn't positve i.e. depth <= 0 , asking the user to re-input a value.
                Console.WriteLine("ERROR: depth must be a positive value! " &
                    "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
            End While

        Next

        Newlines(1)

        ' Q1 part (a).
        ' Part 1 - confirming user's inputted depths.
        For i = 0 To 2
            Console.WriteLine("For distance in " & LCase(Trim(material(i))) & " you entered: " & depth(i) & " metres")
        Next


        Time_Taken_Wood(6) = Calculate_Time_Taken(depth(0), Average_speed(0))
        Time_Taken_Steel(6) = Calculate_Time_Taken(depth(1), Average_speed(1))
        Time_Taken_Titanium(6) = Calculate_Time_Taken(depth(2), Average_speed(2))
        Total_time_Combination = Time_Taken_Wood(6) + Time_Taken_Steel(6) + Time_Taken_Titanium(6)
        Average_Time_taken(6) = Average(Total_time_Combination, 3)

        ' Loop with nested Loop used to enter the variables for each combination of setup into the calculation of a new average time to be assigned into the array.
        For X = 0 To 1
            For I = 0 To 2
                ' Calculates the new average time taken by the effect of modifications, Mod_Average_Speed Array's fist row = Bullet used and Second row = Gunpowder used.
                Mod_Average_Speed_Wood(I, X) = Average_speed(0) * (1 + (Wood_Bullets(I) + Wood_Gunpowder(X)) / 100)
                Mod_Average_Speed_Steel(I, X) = Average_speed(1) * (1 + (Steel_Bullets(I) + Steel_Gunpowder(X)) / 100)
                Mod_Average_Speed_Titanium(I, X) = Average_speed(2) * (1 + (Titanium_Bullets(I) + Titanium_Gunpowder(X)) / 100)
            Next
        Next

        l = 0
        For X = 0 To 1
            For I = 0 To 2
                Time_Taken_Wood(l) = Calculate_Time_Taken(depth(0), Mod_Average_Speed_Wood(I, X))
                Time_Taken_Steel(l) = Calculate_Time_Taken(depth(1), Mod_Average_Speed_Steel(I, X))
                Time_Taken_Titanium(l) = Calculate_Time_Taken(depth(2), Mod_Average_Speed_Titanium(I, X))
                l = l + 1
            Next
        Next

        Newlines(2)
        'producing the tables: 6 tables, Overall Title, Material, Time taken(s)
        For I = 0 To 5

            Console.WriteLine(Space(1) & UCase(Trim(Combination_Titles(I))) & ":")

            Table_Heading(2, "Material", "Time taken (s)", "", "", "")  'Console.Write(Space(2)) at front 
            Console.WriteLine(Space(2) & material(0) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Wood(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(Space(2) & material(1) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Steel(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(Space(2) & material(2) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Titanium(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(StrDup((2 * 17) + 7, "="))

            Total_time_Combination = Time_Taken_Wood(I) + Time_Taken_Steel(I) + Time_Taken_Titanium(I)
            Average_Time_taken(I) = Average(Total_time_Combination, 3)
            Console.WriteLine(Space(1) & "Average time taken: " & Format(Average_Time_taken(I), "0.000") & " seconds")

            Newlines(2)
        Next

        ' Creating a sorted version of the average time array.
        Array.Copy(Average_Time_taken, Sorted_Average_Time_Taken, 7)
        Array.Sort(Sorted_Average_Time_Taken)

        ' Creating the directory and file for the data and writing the headings to the file.
        My.Computer.FileSystem.CreateDirectory("ExperimentalData")
        My.Computer.FileSystem.WriteAllText("ExperimentalData\Data.txt", "=========================================================================" & vbNewLine _
                                            & "  Modifications" & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & "  Average time taken (s)" & vbNewLine _
                                            & "=========================================================================" & vbNewLine, False)

        ' Writing the data to the file.
        For I = 0 To 6
            For P = 0 To 6
                If Sorted_Average_Time_Taken(I) = Average_Time_taken(P) Then
                    My.Computer.FileSystem.WriteAllText("ExperimentalData\Data.txt", Space(2) & Combination_Titles(P) & vbTab & "|" & Space(2) & Format(Sorted_Average_Time_Taken(I), "0.000") & vbTab & vbTab & vbNewLine, True)
                End If
            Next
        Next





    End Sub
End Module