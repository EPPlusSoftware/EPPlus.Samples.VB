Imports EPPlusSamples.ConditionalFormatting
Imports OfficeOpenXml
Imports System

Namespace EPPlusSamples
    Friend Class ConditionalFormattingSamples
        Public Shared Sub Run()
            Console.WriteLine("Running sample 3.2")

            Dim newFile = FileUtil.GetCleanFileInfo("3.2-ConditionalFormattings.xlsx")
            Using pck As ExcelPackage = New ExcelPackage(newFile)
                'Styling in ConditionalFormatting
                StylingExample.Run(pck)

                'Averages in ConditionalFormatting
                'And basic properties for them
                AveragesExample.Run(pck)

                'Standard deviations and top bottom values/percentages
                StandardDeviationTopDown.Run(pck)

                'Last7Days, yesterday,tommorow, last week/month etc.
                DatesAndTime.Run(pck)

                'Format if there are blanks, errors or noBlanks noErrors and duplicate values
                BlanksErrorsAndDuplicates.Run(pck)

                'Formattings that check for text. Cell ends starts and contain given string
                StartsEndContainsExample.Run(pck)

                'Format if equal or between values and custom expressions
                EqualBetweenExpressionsExamples.Run(pck)

                RemovalAndCasting.Run(pck)

                'Advanced CFs below

                'Iconsets rules including custom iconsets
                IconsetsExample.Run(pck)

                'Databars with full features
                DatabarsExample.Run(pck)

                'Two and Three colorscales can use theme, index and auto-color same as databar
                ColorScaleExample.Run(pck)

                pck.SaveAs(newFile)
            End Using

        End Sub
    End Class
End Namespace
