Sub Clear_Output_Data()

'Macro Name: Clear_Output_Data
'Author: Alex Holburn https://www.alexholburn.com
'Last Edit: 03/31/21
'Edit Purpose: Better commenting.
'License: MIT License. Copyright 2021, Alex Holburn https://www.alexholburn.com

'Macro Purpose: The purpose of this macro is to clear the simulation output data with a single button click

'    *******************************BEGIN VARIABLE DECLARATION*******************************

    Dim out_data As Range

'    *************************************BEGIN MAIN CODE************************************
        
        'Set simulation output data range
        Set out_data = Worksheets("Tool_Setup").Range("E6")

        'Write Blanks into the Motion Data Output data ranges
        Worksheets("Motion_Data").Range(out_data).Value = ""
        
        'Write Blanks into the Motion Analysis Data Range
        Worksheets("Motion_Analysis").Range("H9:H21").Value = ""

        'Display a Message to the user so they know the analysis was run.
        MsgBox "Simulation Output Data Cleared!"

End Sub