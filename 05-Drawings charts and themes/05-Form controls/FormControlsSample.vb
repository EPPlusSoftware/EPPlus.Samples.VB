﻿Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports System
Imports System.Drawing
Imports System.Text

Namespace EPPlusSamples.DrawingsChartsAndThemes
    Public Class FormControlsSample
        Public Shared Sub Run()
            Console.WriteLine("Running sample 5.5 - Form controls")
            Using package = New ExcelPackage()
                'First create the sheet containing the data for the check box and the list box.
                Dim dataSheet = CreateDataSheet(package)

                'Create the form-sheet and set headers and som basic properties.
                Dim formSheet = CreateFormSheet(package)

                'Add texts and format the text fields style
                formSheet.Cells("A3").Value = "Name"
                formSheet.Cells("A4").Value = "Gender"
                formSheet.Cells("B3,B5,B11").Style.Border.BorderAround(ExcelBorderStyle.Dotted)
                formSheet.Cells("B3,B5,B11").Style.Fill.SetBackground(eThemeSchemeColor.Background1)

                'Controls are added via the worksheets drawings collection. 
                'Each type has its typed method returning the specific control class. 
                'Optionally you can use the AddControl method specifying the control type via the eControlType enum
                Dim dropDown = formSheet.Drawings.AddDropDownControl("DropDown1")
                dropDown.InputRange = dataSheet.Cells("A1:A2")     'Linkes to the range of items
                dropDown.LinkedCell = formSheet.Cells("G4")        'The cell where the selected index is updated.
                dropDown.SetPosition(3, 1, 1, 0)
                dropDown.SetSize(451, 31)

                formSheet.Cells("A5").Value = "Number of guests"

                'Add a spin button for the number of guests cell
                Dim spinnButton = formSheet.Drawings.AddSpinButtonControl("SpinButton1")
                spinnButton.SetPosition(4, 0, 2, 1)
                spinnButton.SetSize(30, 35)
                spinnButton.Value = 0
                spinnButton.Increment = 1
                spinnButton.MinValue = 0
                spinnButton.MaxValue = 3
                spinnButton.LinkedCell = formSheet.Cells("B5")
                spinnButton.Value = 1

                'Add a group box and four option buttons to select room type
                Dim grpBox = formSheet.Drawings.AddGroupBoxControl("GroupBox 1")
                grpBox.Text = "Room types"
                grpBox.SetPosition(5, 8, 1, 1)
                grpBox.SetSize(150, 150)

                Dim r1 = formSheet.Drawings.AddRadioButtonControl("OptionSingleRoom")
                r1.Text = "Single Room"
                r1.FirstButton = True
                r1.LinkedCell = formSheet.Cells("G7")
                r1.SetPosition(5, 15, 1, 5)

                Dim r2 = formSheet.Drawings.AddRadioButtonControl("OptionDoubleRoom")
                r2.Text = "Double Room"
                r2.LinkedCell = formSheet.Cells("G7")
                r2.SetPosition(6, 15, 1, 5)
                r2.Checked = True

                Dim r3 = formSheet.Drawings.AddRadioButtonControl("OptionSuperiorRoom")
                r3.Text = "Superior"
                r3.LinkedCell = formSheet.Cells("G7")
                r3.SetPosition(7, 15, 1, 5)

                Dim r4 = formSheet.Drawings.AddRadioButtonControl("OptionSuite")
                r4.Text = "Suite"
                r4.LinkedCell = formSheet.Cells("G7")
                r4.SetPosition(8, 15, 1, 5)

                'Group the groupbox together with the radio buttons, so they act as one unit.
                'You can group drawings via the Group method on one of the drawings, here using the group box...
                Dim grp = grpBox.Group(r1, r2, r3)     'This will group the groupbox and three of the radio buttons. You would normally include r4 here as well, but we add it in the next statement to demonstrate how group shapes work.
                '...Or add them to a group drawing returned by the Group method.
                grp.Drawings.Add(r4) 'This will add the fourth radio button to the group

                'Add a scroll bar to control the number of nights
                formSheet.Cells("A11").Value = "Number of nights"
                Dim scrollBar = formSheet.Drawings.AddScrollBarControl("Scrollbar1")
                scrollBar.Horizontal = True    'We want a horizontal scrollbar
                scrollBar.SetPosition(10, 1, 2, 1)
                scrollBar.SetSize(200, 30)
                scrollBar.LinkedCell = formSheet.Cells("B11")
                scrollBar.MinValue = 1
                scrollBar.MaxValue = 365
                scrollBar.Increment = 1
                scrollBar.Page = 7 'How much a page click should increase.
                scrollBar.Value = 1

                'Add a listbox and connect it to the input range in the data sheet
                formSheet.Cells("A12").Value = "Requests"
                Dim listBox = formSheet.Drawings.AddListBoxControl("Listbox1")
                listBox.InputRange = dataSheet.Cells("B1:B3")
                listBox.LinkedCell = formSheet.Cells("G12")
                listBox.SetPosition(11, 5, 1, 0)
                listBox.SetSize(200, 100)

                'Last, add a button and connect it to a macro appending the data to a text file.
                Dim button = formSheet.Drawings.AddButtonControl("ExportButton")
                button.Text = "Make Reservation"
                button.Macro = "ExportButton_Click"
                button.SetPosition(15, 0, 1, 0)
                button.AutomaticSize = True
                formSheet.Select(formSheet.Cells("B3"))

                package.Workbook.CreateVBAProject()
                Dim [module] = package.Workbook.VbaProject.Modules.AddModule("ControlEvents")
                Dim code = New StringBuilder()
                code.AppendLine("Sub ExportButton_Click")
                code.AppendLine("Msgbox ""Here you can place the code to handle the form""")
                code.AppendLine("End Sub")
                [module].Code = code.ToString()

                package.SaveAs(FileUtil.GetCleanFileInfo("5.5-FormControls.xlsm"))
                Console.WriteLine("Sample 5.5 finished.")
                Console.WriteLine()
            End Using
        End Sub

        Private Shared Function CreateFormSheet(ByVal package As ExcelPackage) As ExcelWorksheet
            Dim formSheet = package.Workbook.Worksheets.Add("Form")
            formSheet.Cells("A1").Value = "Room booking"
            formSheet.Cells("A1").Style.Font.Size = 18
            formSheet.Cells("A1").Style.Font.Bold = True
            formSheet.Columns(1).Width = 30
            formSheet.Columns(2).Width = 60
            formSheet.Cells.Style.Fill.SetBackground(Color.Gray)

            formSheet.Rows(1, 18).Height = 25

            Return formSheet
        End Function

        Private Shared Function CreateDataSheet(ByVal package As ExcelPackage) As ExcelWorksheet
            Dim dataSheet = package.Workbook.Worksheets.Add("Data")
            dataSheet.Cells("A1").Value = "Man"
            dataSheet.Cells("A2").Value = "Woman"

            dataSheet.Cells("B1").Value = "Garden view"
            dataSheet.Cells("B2").Value = "Sea view"
            dataSheet.Cells("B3").Value = "Parking lot view"

            dataSheet.Hidden = eWorkSheetHidden.Hidden 'We hide the data sheet.

            Return dataSheet
        End Function
    End Class
End Namespace
