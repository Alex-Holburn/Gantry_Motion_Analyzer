Sub Analyze_Motion()

'Macro Name: Analyze_Motion
'Author: Alex Holburn https://www.alexholburn.com
'Last Edit: 03/31/21
'Edit Purpose: Better commenting.
'License: MIT License. Copyright 2021, Alex Holburn https://www.alexholburn.com

'Macro Purpose: The purpose of this macro is to rapidly compare what the cycle times for non-coordinated vs linear coordinated motion
'               so one can rapidly determine if the time savings from coordinated motion is worth the programming effort or purchasing
'               different hardware is the only fix.

'    *******************************BEGIN VARIABLE DECLARATION*******************************

    Dim min_xaxis_distance As Variant    'Minimum x-axis Distance Traveled
    Dim min_yaxis_distance As Variant    'Minimum y-axis Distance Traveled
    Dim min_zaxis_distance As Variant    'Minimum z-axis Distance Traveled

    Dim max_xaxis_distance As Variant    'Maximum x-axis Distance Traveled
    Dim max_yaxis_distance As Variant    'Maximum y-axis Distance Traveled
    Dim max_zaxis_distance As Variant    'Maximum z-axis Distance Traveled
    
    Dim avg_xaxis_distance As Variant    'Average x-axis Distance Traveled
    Dim avg_yaxis_distance As Variant    'Average y-axis Distance Traveled
    Dim avg_zaxis_distance As Variant    'Average z-axis Distance Traveled
    
    Dim max_xaxis_velocity As Variant    'Maximum x-axis velocity
    Dim max_yaxis_velocity As Variant    'Maximum y-axis velocity
    Dim max_zaxis_velocity As Variant    'Maximum z-axis velocity
    
    Dim min_xaxis_time As Variant        'Minimum x-axis Travel Time
    Dim min_yaxis_time As Variant        'Minimum y-axis Travel Time
    Dim min_zaxis_time As Variant        'Minimum z-axis Travel Time
    Dim min_total_time As Variant        'Minimum total Travel Time
    
    Dim max_xaxis_time As Variant        'Minimum x-axis Travel Time
    Dim max_yaxis_time As Variant        'Minimum y-axis Travel Time
    Dim max_zaxis_time As Variant        'Minimum z-axis Travel Time
    Dim max_total_time As Variant        'Maximum total Travel Time
    
    Dim avg_xaxis_time As Variant        'Average x-axis Travel Time
    Dim avg_yaxis_time As Variant        'Average y-axis Travel Time
    Dim avg_zaxis_time As Variant        'Average z-axis Travel Time
    Dim avg_total_time As Variant        'Maximum total Travel Time
    
    Dim sim_xaxis_step As Variant        'Step size of the x-axis simulation range
    Dim sim_yaxis_step As Variant        'Step size of the y-axis simulation range
    Dim sim_zaxis_step As Variant        'Step size of the z-axis simulation range
    
    Dim xaxis_time As Variant            'Simulation Placeholder for x-axis travel time.
    Dim yaxis_time As Variant            'Simulation Placeholder for y-axis travel time.
    Dim zaxis_time As Variant            'Simulation Placeholder for z-axis travel time.
    
    Dim number_of_points As Variant      'Number of simulation points to generate for external software.
    
    Dim xaxis_traveldata_range As Range  'Range of Travel Time Data of x-axis for statistical Analysis.
    Dim yaxis_traveldata_range As Range  'Range of Travel Data of y-axis for statistical Analysis.
    Dim zaxis_traveldata_range As Range  'Range of Travel Data of z-axis for statistical Analysis.
    
    Dim Output_Template As Variant       'Path to the output data template
    Dim Output_Filename As Variant       'Name of the output data file
    Dim out_data As Range                'Output Motion Data Range
    

'    *************************************BEGIN MAIN CODE************************************
    
        'Set The Minimum Axis Travel Distances
        min_xaxis_distance = Sheets("Motion_Analysis").Range("B9").Value
        min_yaxis_distance = Sheets("Motion_Analysis").Range("B13").Value
        min_zaxis_distance = Sheets("Motion_Analysis").Range("B17").Value
        
        'Set The Maximum Axis Travel Distances
        max_xaxis_distance = Sheets("Motion_Analysis").Range("B10").Value
        max_yaxis_distance = Sheets("Motion_Analysis").Range("B14").Value
        max_zaxis_distance = Sheets("Motion_Analysis").Range("B18").Value
        
        'Set The Maximum Axis Velocities
        max_xaxis_velocity = Sheets("Motion_Analysis").Range("B12").Value
        max_yaxis_velocity = Sheets("Motion_Analysis").Range("B16").Value
        max_zaxis_velocity = Sheets("Motion_Analysis").Range("B20").Value
        
        'Set The Number of Simulation Points
        number_of_points = Sheets("Motion_Analysis").Range("B22").Value
        
        'Determine the step sizes for all of the axis.
        sim_xaxis_step = (max_xaxis_distance - min_xaxis_distance) / (number_of_points - 1)
        sim_yaxis_step = (max_yaxis_distance - min_yaxis_distance) / (number_of_points - 1)
        sim_zaxis_step = (max_zaxis_distance - min_zaxis_distance) / (number_of_points - 1)
        
        'Set Travel Time Data Ranges for all axis.
        Set xaxis_traveldata_range = Worksheets("Motion_Data").Range("B:B")
        Set yaxis_traveldata_range = Worksheets("Motion_Data").Range("C:C")
        Set zaxis_traveldata_range = Worksheets("Motion_Data").Range("D:D")
        
        'Set the workbook output paths
        Output_Template = Worksheets("Tool_Setup").Range("H6")
        Output_Filename = Worksheets("Tool_Setup").Range("H5")

        'Set simulation output data range
        Set out_data = Worksheets("Tool_Setup").Range("E6")
            
        'Determine which algorithm to use for move data generation and run it
        If Sheets("Motion_Analysis").Range("B21").Value = "NO" Then 'Single-Axis Moves Algorithm

                'We run a for loop here for the number of simulated points across the gantry travel.
                For i = 0 To number_of_points - 1
                
                'Determine Axis Travel Times Per Simulation Step
                xaxis_time = ((min_xaxis_distance + (sim_xaxis_step * i)) / max_xaxis_velocity)
                yaxis_time = ((min_yaxis_distance + (sim_yaxis_step * i)) / max_yaxis_velocity)
                zaxis_time = ((min_zaxis_distance + (sim_zaxis_step * i)) / max_zaxis_velocity)
                
                'Write the simulation number for external software processing.
                Sheets("Motion_Data").Range("A1").Offset(1 + i, 0) = i + 1
                
                'Calculate the travel times of each axis
                Sheets("Motion_Data").Range("B1").Offset(1 + i, 0) = xaxis_time
                Sheets("Motion_Data").Range("C1").Offset(1 + i, 0) = yaxis_time
                Sheets("Motion_Data").Range("D1").Offset(1 + i, 0) = zaxis_time
                
                'Determine the simulated position in meters for all axis.
                Sheets("Motion_Data").Range("E1").Offset(1 + i, 0) = min_xaxis_distance + (sim_xaxis_step * i)
                Sheets("Motion_Data").Range("F1").Offset(1 + i, 0) = min_yaxis_distance + (sim_yaxis_step * i)
                Sheets("Motion_Data").Range("G1").Offset(1 + i, 0) = min_zaxis_distance + (sim_zaxis_step * i)

                'Write the x-axis velocity data
                Sheets("Motion_Data").Range("H1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("E1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("B1").Offset(1 + i, 0)
                    
                'Write the y-axis velocity data
                Sheets("Motion_Data").Range("I1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("F1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("C1").Offset(1 + i, 0)
                    
                'Write the y-axis velocity data
                Sheets("Motion_Data").Range("J1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("G1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("D1").Offset(1 + i, 0)
                
                Next i
            
         Else: 'Coordinated Motion Algorithm
            
                'We run a for loop here for the number of simulated points across the gantry travel.
                For i = 0 To number_of_points - 1
                
                'Determine Axis Travel Times Per Simulation Step
                xaxis_time = ((min_xaxis_distance + (sim_xaxis_step * i)) / max_xaxis_velocity)
                yaxis_time = ((min_yaxis_distance + (sim_yaxis_step * i)) / max_yaxis_velocity)
                zaxis_time = ((min_zaxis_distance + (sim_zaxis_step * i)) / max_zaxis_velocity)
                
                'Write the simulation number for external software processing.
                Sheets("Motion_Data").Range("A1").Offset(1 + i, 0) = i + 1
                
                'Determine the simulated position in meters for all axis.
                Sheets("Motion_Data").Range("E1").Offset(1 + i, 0) = min_xaxis_distance + (sim_xaxis_step * i)
                Sheets("Motion_Data").Range("F1").Offset(1 + i, 0) = min_yaxis_distance + (sim_yaxis_step * i)
                Sheets("Motion_Data").Range("G1").Offset(1 + i, 0) = min_zaxis_distance + (sim_zaxis_step * i)
                
                    'Determine if x-axis is slowest, and clamp other axis speeds to match move time.
                    If xaxis_time >= yaxis_time And xaxis_time >= zaxis_time Then
                        
                        Sheets("Motion_Data").Range("B1").Offset(1 + i, 0) = xaxis_time
                        Sheets("Motion_Data").Range("C1").Offset(1 + i, 0) = xaxis_time
                        Sheets("Motion_Data").Range("D1").Offset(1 + i, 0) = xaxis_time
                                
                    'Determine if y-axis is slowest, and clamp other axis speeds to match move time.
                    ElseIf yaxis_time >= xaxis_time And yaxis_time >= zaxis_time Then

                        Sheets("Motion_Data").Range("B1").Offset(1 + i, 0) = yaxis_time
                        Sheets("Motion_Data").Range("C1").Offset(1 + i, 0) = yaxis_time
                        Sheets("Motion_Data").Range("D1").Offset(1 + i, 0) = yaxis_time
                    
                    'Determine if z-axis is slowest, and clamp other axis speeds to match move time.
                    ElseIf zaxis_time >= xaxis_time And zaxis_time >= yaxis_time Then

                        Sheets("Motion_Data").Range("B1").Offset(1 + i, 0) = yaxis_time
                        Sheets("Motion_Data").Range("C1").Offset(1 + i, 0) = yaxis_time
                        Sheets("Motion_Data").Range("D1").Offset(1 + i, 0) = yaxis_time
                    
                    End If
                    
                    'Write the x-axis velocity data
                    Sheets("Motion_Data").Range("H1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("E1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("B1").Offset(1 + i, 0)
                    
                    'Write the y-axis velocity data
                    Sheets("Motion_Data").Range("I1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("F1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("C1").Offset(1 + i, 0)
                    
                    'Write the y-axis velocity data
                    Sheets("Motion_Data").Range("J1").Offset(1 + i, 0) = Sheets("Motion_Data").Range("G1").Offset(1 + i, 0) / Sheets("Motion_Data").Range("D1").Offset(1 + i, 0)
                
                Next i
         
         End If
         
         'Determine the minimum axis travel times
         min_xaxis_time = Application.WorksheetFunction.Min(xaxis_traveldata_range)
         min_yaxis_time = Application.WorksheetFunction.Min(yaxis_traveldata_range)
         min_zaxis_time = Application.WorksheetFunction.Min(zaxis_traveldata_range)

         'Determine the maximum axis travel times
         max_xaxis_time = Application.WorksheetFunction.Max(xaxis_traveldata_range)
         max_yaxis_time = Application.WorksheetFunction.Max(yaxis_traveldata_range)
         max_zaxis_time = Application.WorksheetFunction.Max(zaxis_traveldata_range)

         'Determine the mean (average) axis travel times
         avg_xaxis_time = Application.WorksheetFunction.Average(xaxis_traveldata_range)
         avg_yaxis_time = Application.WorksheetFunction.Average(yaxis_traveldata_range)
         avg_zaxis_time = Application.WorksheetFunction.Average(zaxis_traveldata_range)

        'Determine which algorithm to use for determining total move data and run it
        If Sheets("Motion_Analysis").Range("B21").Value = "NO" Then 'Single-Axis Moves Algorithm
            
            min_total_time = min_xaxis_time + min_yaxis_time + min_zaxis_time
            max_total_time = max_xaxis_time + max_yaxis_time + max_zaxis_time
            avg_total_time = avg_xaxis_time + avg_yaxis_time + avg_zaxis_time
            
        'Coordinated Moves Algorithm. We check to make sure all travel values are equal.
        ElseIf avg_xaxis_time = avg_yaxis_time And avg_xaxis_time = avg_zaxis_time Then
        
            min_total_time = min_xaxis_time
            max_total_time = max_xaxis_time
            avg_total_time = avg_xaxis_time
            
        Else: 'Write an error if the axis values do not match when using coordinated motion.

            min_total_time = "ERROR"
            max_total_time = "ERROR"
            avg_total_time = "ERROR"
                 
        End If
         
         'Write the x-axis Output Data
         
         Sheets("Motion_Analysis").Range("H9").Value = min_xaxis_time
         Sheets("Motion_Analysis").Range("H10").Value = max_xaxis_time
         Sheets("Motion_Analysis").Range("H11").Value = avg_xaxis_time
         
         'Write the y-axis Output Data
         
         Sheets("Motion_Analysis").Range("H12").Value = min_yaxis_time
         Sheets("Motion_Analysis").Range("H13").Value = max_yaxis_time
         Sheets("Motion_Analysis").Range("H14").Value = avg_yaxis_time
         
         'Write the z-axis Output Data
         
         Sheets("Motion_Analysis").Range("H15").Value = min_zaxis_time
         Sheets("Motion_Analysis").Range("H16").Value = max_zaxis_time
         Sheets("Motion_Analysis").Range("H17").Value = avg_zaxis_time
         
        'Write the Total Move Output Data
         
         Sheets("Motion_Analysis").Range("H18").Value = min_total_time
         Sheets("Motion_Analysis").Range("H19").Value = max_total_time
         Sheets("Motion_Analysis").Range("H20").Value = avg_total_time
         
         'Write the Type of Motion Analyzed
         If Sheets("Motion_Analysis").Range("B21").Value = "NO" Then
         
            Sheets("Motion_Analysis").Range("H21").Value = "Single-Axis Moves"
        
         Else:
            
            Sheets("Motion_Analysis").Range("H21").Value = "Coordinated Motion"
        
         End If
         
         'Write Output Data to file if required.
         If Sheets("Motion_Analysis").Range("B23").Value = "YES" Then
            'Open the template file for editing
            Workbooks.Add (Output_Template)
            
            'Copy the output data to the template workbook.
            Workbooks(1).Worksheets("Motion_Data").Range(out_data).Copy _
                Workbooks(2).Worksheets("Sheet1").Range("A2")
                
            'Save the template workbook as the output filename.
            Workbooks(2).SaveAs (Output_Filename)
            
            'Close the output data workbook
            Workbooks(2).Close
         
         End If
         
         'Determine which Process Complete Message to Display.
         If Sheets("Motion_Analysis").Range("B23").Value = "NO" Then
            
            'Display a Message to the user so they know the analysis was run.
            MsgBox "Analysis Complete!"
         
         Else:
         
            'Display a Message to the user so they know the analysis was run and data save to a file.
            MsgBox "Analysis Complete!" & vbCrLf & vbCrLf & "Output File Saved As:" & vbCrLf & vbCrLf & Sheets("Tool_Setup").Range("H5")
         
         End If
         

    
                
End Sub