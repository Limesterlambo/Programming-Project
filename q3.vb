'Final edit at 16:20 on 14/12/2018
'The program allows the user to evaluate the average time taken for a bullet to travel through three distances for each material
'(wood, steel, titanium). The program averages the distances and prints out the average time of travel through each material;
'this data can be used for the purpose of figuring out which is the safest material for testing purposes.
'Main Contributor: Oliver Monk
Module Module1

    Function Newlines(Number_of_Spaces)
        For i = 1 To Number_of_Spaces
            Console.Write(vbNewLine)
        Next
    End Function
'Function creates a line break

    Sub Main()
        Dim D1A, D1B, D1C As Double : Dim D2A, D2B, D2C As Double : Dim D3A, D3B, D3C As Double
        'Creating Dim names for three distances of the three materials
        Dim Wood_AD, Steel_AD, Titanium_AD As Double
        'Creating Dim names for average distances of the three materials
        Dim Wood_AT, Steel_AT, Titanium_AT As Double
        'Creating Dim names for average bullet times through the three materials
        Dim Velocity_Wood, Velocity_Steel, Velocity_Titanium As Double
        'Creating Dim names for optimal bullet velocity in each material
        Dim OVPI_Wood, OVPI_Steel, OVPI_Titanium As Double
        'Creating Dim names for Optimal Velocity Percentage Increase values 

        Console.WriteLine("This program will output the average time taken for a bullet to travel through")
        Console.WriteLine("the three materials, when given three different depth values. It utilises the optimal")
        Console.WriteLine("weapon attachment setup for each material with regard to bullet velocity.")
        Newlines(1)
        Console.WriteLine("The purpose of this program is to determine whether or not there is a material distance")
        Console.WriteLine("scenario wherein the optimal weapon setup is no longer optimal.")
        Newlines(1)
        'Introducing the program to the user

        Velocity_Wood = 323.5
        Velocity_Steel = 180.8
        Velocity_Titanium = 100.3
        'Defining standard bullet velocity in each material

        OVPI_Wood = 55
        OVPI_Steel = 20
        OVPI_Titanium = 35
        'Defining optimal bullet velocity percentage increases for each material

        Console.WriteLine("Please enter three distance (depth of material) values for each material.")
        Newlines(1)
        'Asks user to input values
        Console.Write("Wood, Distance 1: ") : D1A = Console.ReadLine()
        Console.Write("Wood, Distance 2: ") : D1B = Console.ReadLine()
        Console.Write("Wood, Distance 3: ") : D1C = Console.ReadLine()
        Newlines(1)
        'Gathers distance values for Wood
        Console.Write("Steel, Distance 1: ") : D2A = Console.ReadLine()
        Console.Write("Steel, Distance 2: ") : D2B = Console.ReadLine()
        Console.Write("Steel, Distance 3: ") : D2C = Console.ReadLine()
        Newlines(1)
        'Gathers distance values for Steel
        Console.Write("Titanium, Distance 1: ") : D3A = Console.ReadLine()
        Console.Write("Titanium, Distance 2: ") : D3B = Console.ReadLine()
        Console.Write("Titanium, Distance 3: ") : D3C = Console.ReadLine()
        Newlines(1)
        'Gathers distance values for Titanium

        Wood_AD = (D1A + D1B + D1C) / 3
        Steel_AD = (D2A + D2B + D2C) / 3
        Titanium_AD = (D3A + D3B + D3C) / 3
        'Calculates average distance for each material

        Wood_AT = Wood_AD / (Velocity_Wood + (Velocity_Wood * (OVPI_Wood / 100)))
        Steel_AT = Steel_AD / (Velocity_Steel + (Velocity_Steel * (OVPI_Steel / 100)))
        Titanium_AT = Titanium_AD / (Velocity_Titanium + (Velocity_Titanium * (OVPI_Titanium / 100)))
        'Calculated average time for each material

        Console.WriteLine("Average time taken for bullet to pass through the average wood distance is: " & Format(Wood_AT, "0.000") & "seconds")
        Console.WriteLine("Average time taken for bullet to pass through the average steel distance is: " & Format(Steel_AT, "0.000") & "seconds")
        Console.WriteLine("Average time taken for bullet to pass through the average titanium distance is: " & Format(Titanium_AT, "0.000") & "seconds")
        Newlines(1)
        'Prints out average time values to user

        Console.WriteLine("It has been found that the optimal setup remains constant regardless of material distance.")
        Console.WriteLine("This is not what one would expect in a true-to-life scenario. One potential reason for this")
        Console.WriteLine("is that a friction value (co-efficient) is not specified for each material, therefore there")
        Console.WriteLine("can be no decelleration experienced by the bullet. Another reason why the optimal setup")
        Console.WriteLine("remains constant is that the angle of penetration, and thus the bullet's direction of travel")
        Console.WriteLine("is assumed to be 90 degrees relative to the material. This means that the bullet maintains a")
        Console.WriteLine("constant horizontal flight path, and no gravitational acceleration is experienced.")
        Newlines(1)
        'Discusses the findings of the program

    End Sub

End Module

